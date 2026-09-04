# Throttling: what bites, and how we stop being blocked

Status: research and design. Nothing in this document is implemented yet; it is the plan we
execute from. Written 2026-09-04 against v0.5.1 (`origin/main` at `ed69a4a`).

## 1. Symptoms and measured evidence

### 1.1 What the operator sees

Since late August every long day of use ends the same way: a `--mention` post refuses, an
attachment read refuses, or the inbox goes quiet for an hour. The failures are not random —
they cluster on the endpoints we call *least*, which is the first clue about where the budget
is actually being spent.

**2026-09-04, CTP daemon** (poll interval 60 s, one signed-in user, three allowlisted chats).
Every `--mention` post between ~07:45Z and ~08:05Z failed with

```
THROTTLED: the member list refresh for mention resolution was throttled; nothing was done
  ... retry after 62s
```

Four attempts 70 s apart, then further attempts at 180 s spacing — all refused. **Untagged posts
into the same chat succeeded throughout the same window.** So this was never an account-wide or
tenant-wide block: `POST /chats/{id}/messages` was healthy while `GET /chats/{id}/members` was
not.

**Earlier in the same week**, two other shapes of the same failure:

- `/me` calls starved the same budget: every fresh CLI process paid, and lost, the same throttled
  `GET /me` call, so `teams-send-file` refused one process at a time (KNOWN-ISSUES, "send_chat_file:
  a throttled /me refused every CLI attempt"). Fixed in 0.5.1 by the persisted self-id cache —
  which is the proof that *removing a call class entirely* is what works.
- Ad-hoc single-message GETs answered 429 with `retry-after: 62` on **every** attempt across
  20+ minutes of patient, Retry-After-honouring backoff, while the poller ran (README,
  "Downloading attachments"). Waiting was measured dead; only the poller going quiet made room,
  which is why `src/inbox-yield.ts` exists.

### 1.2 The 2026-09-04 incident, root-caused

The members cache on the CTP instance
(`~/.teams-assistant-guidewire/.members-cache.json`) holds:

| chat | roster `fetchedAt` | members |
| --- | --- | --- |
| `19:1120cc7c…` — CTP agent-team | 2026-09-03T07:34:05Z | 6 |
| `19:8af48977…` — MCP dev test | 2026-09-01T11:42:20Z | 3 |
| `19:7cfde672…` — Guidewire Management | **no entry, ever** | — |

`DEFAULT_MEMBERS_TTL_MS` is 24 h. The CTP agent-team roster therefore expired at
**2026-09-04T07:34:05Z**. The first throttled mention was ~07:45Z — **eleven minutes after the
TTL ran out.** That is not a coincidence; it is the mechanism:

1. TTL expiry turns `MembersCache.get()` into a miss (`get()` returns `undefined` past TTL —
   there is no stale-serve path).
2. A miss makes `resolveMentions` call `refreshMembers`, i.e. the one endpoint we know is
   contended.
3. The refresh 429s, so `cacheIfNonEmpty` is never reached — **the cache is never rewritten**.
4. Every subsequent attempt repeats steps 1–3 forever. The cache cannot heal itself, because
   healing requires the very call that is being refused.

So the 24 h TTL does not bound staleness risk at the cost of one refresh; under contention it
converts a working cached path into a **permanent** hard dependency on the throttled endpoint.
The third chat is worse still: it has no entry at all, so *every* mention into the Guidewire
Management chat has always been a live `/members` call.

### 1.3 Our actual call volume — and why it cannot be the cause

Both daemons on this host run the same client id
`d3590ed6-52b3-4102-aeff-aad2292ab01c` (the Microsoft **Office** first-party application id),
signed in as two different users in two different tenants.

`readMessages` issues exactly one `GET /chats/{id}/messages?$top=50` per chat per cycle
(the watermark is applied client-side in `applyWatermark`, so the request is identical whether
or not anything is new). Steady-state:

| caller | requests | per day |
| --- | --- | --- |
| CTP poller — 3 chats @ 60 s | 3 / min | 4 320 |
| SRP poller — 4 chats @ 60 s | 4 / min | 5 760 |
| `/me` | ~0 since 0.5.1 (persisted cache) | ~0 |
| `/chats/{id}/members` | 1 per mention on a cache miss | single digits |
| single-message GET | 1 per quoted reply / attachment read | single digits |
| attachment download | 1 message GET + 1 `/shares` + 1 content GET per file | single digits |

**Per mailbox that is 30 requests per 10 minutes.** Whatever budget we are exhausting, we are
not exhausting it with volume — three requests a minute is not a load. This is the single most
important measurement in this document, because it rules out "poll less often" as a *fix* and
points at *which bucket the key is scoped to* as the real question (section 2).

## 2. Which limit actually applies

### 2.1 Teams throttling has four independent keys, and we only reasoned about one

Microsoft's own words on the Teams service limits
([throttling-limits](https://learn.microsoft.com/en-us/graph/throttling-limits)):

> Microsoft Teams applies throttling limits across four independent dimensions. A request is
> throttled when it exceeds any limit that applies to it, so always design for the lowest limit
> that your scenario hits.

| dimension | what it counts |
| --- | --- |
| **Per app** | all requests from one app (client id) summed **across every tenant** |
| **Per app per tenant** | one app's requests within a single tenant — "the most commonly reached limit" |
| **Per resource** | requests against a single team, channel or chat |
| **Per user** | requests on behalf of a single user (the delegated case) |

> Each limit is evaluated over a short burst window. A sustained limit of approximately 83 percent
> of the listed value is also evaluated over a longer window, so a workload that runs continuously
> at the listed rate can still be throttled.

Our README and `SETUP.md` both state the budget is "keyed by client id *and* signed-in user
together." **That is one of the four keys, and not the one that is biting us.** The document you
are reading exists largely to correct that.

### 2.2 The endpoints we call, mapped to their rows

| our call | per app | per app per tenant | per resource | per user |
| --- | --- | --- | --- | --- |
| `GET /chats/{id}/messages` (the poller) | 200 rps | 20 rps | 1 rps / chat | – |
| `POST /chats/{id}/messages` (every post) | 200 rps | 20 rps | 1 rps / chat | 1 rps |
| `GET …/hostedContents` (attachments) | 500 rps | 50 rps | 1 rps / chat | – |
| `GET …/hostedContents/{id}/$value` | 600 rps | 60 rps | 1 rps / chat | – |
| **`GET /chats/{id}/members`** | **1 500 rps** | **30 rps** | **1 rps / chat** | **1 rps** |

The members row is the finding. **`GET /chats/{chat-id}/members` is not itemised anywhere in the
Teams limits tables** — the Members table lists only `GET /teams/{team-id}/members` (teams, not
chats) plus `POST`/`DELETE` for chat members. By the page's own fallback rule —

> Any Microsoft Teams request that isn't listed in the preceding tables uses these [Default limits].

— our member-list read lands in the **generic default-GET bucket**. That bucket is not a private
allowance for us. It is the shared catch-all for *every un-itemised Teams GET made by our client
id*, and it is the tightest per-tenant number on the page (30 rps) after the per-user 1 rps.

`GET /me` is a third mechanism again: the **Identity and access service limits** section governs
the `user` resource with a Resource-Unit token bucket keyed by application, tenant, and
application+tenant pair — not the Teams RPS buckets at all. Which is why a throttled `/me` and a
throttled `/members` were two different incidents with two different fixes.

`chatMessage`, `chatMessageHostedContent` and `conversationMember` are all explicitly listed as
Teams-section resources; none of them appear in the Outlook section's resource list (`message`,
`mailFolder`, `event`, `calendar`, `contact`, …). **The Outlook 10 000-requests-per-10-minutes
per-app-per-mailbox budget is therefore almost certainly not the budget we are exhausting.** The
docs never say so in prose — this is inference from the disjoint resource-list partitioning
across the Identity, Teams and Outlook sections — but it fits the evidence: we run at 30 requests
per 10 minutes per mailbox, 0.3 % of the Outlook allowance, and still get refused.

### 2.3 The limit that actually bites

The "Per app" dimension is defined by Microsoft as:

> **Per app** — All requests from one app (client id) summed across every tenant. Multitenant apps
> that serve many customers.

**The bucket key is literally the client id, with no distinction for who owns the app.** Both our
daemons authenticate as `d3590ed6-52b3-4102-aeff-aad2292ab01c` — the Microsoft **Office**
first-party application id. We chose it because it needs no app registration and no admin consent;
that convenience is exactly what put us in someone else's queue.

The consequence, stated plainly:

- **Per app per tenant (30 rps, default GETs):** inside If P&C's tenant, every other caller
  presenting the Office client id draws on the *same* 30 rps allowance we do. We contribute three
  requests a minute to a pool sized for a whole tenant's Office estate. We do not control it, we
  cannot see it, and we cannot make it quieter.
- **Per app (1 500 rps, global):** pooled across every tenant on earth using that client id.

That is why the throttle looks the way it does: **untagged posts succeed while `--mention` fails.**
`POST /chats/{id}/messages` is *itemised*, with its own 20 rps-per-tenant row — a bucket only
actual chat posts draw on. `GET /chats/{id}/members` is *un-itemised*, sharing the crowded default
bucket. Our volume never changed between the two calls; the bucket did.

**So the answer to "which limit bites" is: the Teams default-GET bucket for un-itemised reads,
keyed per app per tenant, on a client id we share with an entire customer estate.** Every
mitigation below is graded first on whether it changes that key or removes calls from that bucket.

### 2.4 Retry-After, and why waiting was measured dead

> Backing off requests using the Retry-After delay is the fastest way to recover from throttling
> **because Microsoft Graph continues to log resource usage while a client is being throttled.**
> … Avoid immediate retries, because all requests accrue against your usage limits.

Our gate discipline (0.2.0–0.4.1) is correct and should stay. But the 2026-09-02 measurement —
429 on every attempt across 20+ minutes of honest backoff — is explained by this section: if the
bucket is drained by *other people in the customer's tenant*, our silence does not refill it.
Retry-After tells us when *Graph* will re-evaluate; it does not promise the shared pool has room.
This is the structural reason quota-yield (`src/inbox-yield.ts`) helped only partially: it
rations *our* share of a pool we are not the main consumer of.

## 3. Options

Graded first on the question section 2 makes decisive: **does it change the bucket key, or take
calls out of the crowded default-GET bucket?** Anything that only makes us tidier inside a bucket
other people are draining scores low, however good the engineering.

| # | option | call-rate effect | admin / consent needed | attribution | effort | risk |
| --- | --- | --- | --- | --- | --- | --- |
| 1 | Log `x-ms-throttle-scope` on every 429 | none (diagnostic) | none | unchanged | ~2 h | none — this is how we stop guessing |
| 2 | Roster harvested from poll results + stale-serve | `/members` → **~0**; removes mentions from the default-GET bucket entirely | none | unchanged | ~1 day | roster limited to people who have spoken; stale entries |
| 3 | Split the two daemons onto different client ids | no change in count; splits the `Tenant_Application` key | none | unchanged | env var only | still a shared first-party id, just a different queue |
| 4 | Longer poll interval (60 s → 180 s) | polls 10 080 → 3 360 / day | none | unchanged | env var only | inbox latency; does not address the real key |
| 5 | `$filter`/`$orderby` on the poll | same call count, smaller payloads | none | unchanged | ~half a day | no documented throttle discount |
| 6 | `$batch` the polls | **none — rejected** | – | – | – | Graph charges every inner request separately |
| 7 | One shared poller for both instances | **none — rejected** | – | – | – | two users in two tenants; the buckets are already separate |
| 8 | **Own multi-tenant app registration (delegated)** | no change in count; **own `Tenant_Application` + `Application` buckets** | user consent likely enough; publisher verification needed | **preserved — posts stay from the named human** | ~1 week | step-up consent; CA may block device-code flow |
| 9 | Change notifications → replace polling | polls 10 080 → **~0/day** + backstop | none beyond option 8's scopes | unchanged | ~2 weeks + Azure infra | subscription renewal; delivery reliability |
| 10 | Teams bot for ingestion only | polls → **0**; documented, generous bot limits | Teams app install (admin approval likely) | reads only — **posts stay delegated** | ~3 weeks | very visible to tenant admins |
| 11 | Teams bot for posting | – | Teams app install | **fails — posts appear as the bot** | – | disqualified by the attribution rule |

### 3.1 Notes that decide the table

**Option 2 — the roster we are already being handed for free.** `toChatMessage` already maps
Graph's sender id onto `ChatMessage.fromId`, so *every polled message hands us an
(AAD id, displayName) pair at zero marginal Graph cost.* Those are exactly the pairs
`resolveMentionTargets` needs, and they cover precisely the population worth @mentioning: people
who actually talk in the chat. Combine with (a) serving a stale roster when a refresh 429s instead
of failing the send, and (b) treating TTL expiry as "refresh when convenient", never as
"this cache is now empty" — and the 2026-09-04 incident cannot recur, because the send path stops
depending on `/members` at all.

**Option 6 is not a matter of degree — Microsoft is explicit.** From the batching guidance:
"Requests in a batch are evaluated individually against the applicable throttling limits and if
any request exceeds the limits, it fails with a status of 429." Batching saves round-trips, not
budget. ([json-batching](https://learn.microsoft.com/en-us/graph/json-batching),
[throttling](https://learn.microsoft.com/en-us/graph/throttling))

**Option 7 dies on section 2's key structure.** The two daemons sign in as different users in
different tenants; their `Tenant_Application` buckets are already distinct. Merging the pollers
would issue the same seven requests a minute from one process instead of two. There is nothing
to win.

**Delta query is closed to us.** `GET /users/{id}/chats/getAllMessages/delta` lists delegated
permissions as **"Not supported"** — application permission only
([chatMessage: delta](https://learn.microsoft.com/en-us/graph/api/chatmessage-delta?view=graph-rest-1.0)).
A delegated daemon cannot use it. Option 5 (`$filter=lastModifiedDateTime gt …` with a matching
`$orderby`, documented on
[List messages in a chat](https://learn.microsoft.com/en-us/graph/api/chat-list-messages?view=graph-rest-1.0))
is the delegated-compatible shadow of it, and it is a payload optimisation, not a budget one.

**Option 8, the consent detail that shapes everything.** Delegated permissions, from the
[permissions reference](https://learn.microsoft.com/en-us/graph/permissions-reference):

| permission | admin consent required |
| --- | --- |
| `Chat.Read` | **No** |
| `Chat.ReadWrite` | **No** |
| `ChatMessage.Send` | **No** |
| `User.Read` | **No** |
| `User.ReadBasic.All` | **No** |
| **`ChatMember.Read`** | **Yes** |
| **`ChatMember.ReadWrite`** | **Yes** |

So an own app registration that asks only for `Chat.ReadWrite`, `ChatMessage.Send` and `User.Read`
needs **no Global Admin in If's tenant** under a default consent policy — the consultant consents
for himself, once. The moment we also ask for `ChatMember.Read` we need a tenant admin, and the
whole "no admin rights required" property collapses.

**That makes option 2 a prerequisite for option 8, not merely a companion to it.** Under the
user-consentable scope set we would have no `/chats/{id}/members` access at all — the roster
harvested from message senders is then the *only* way mentions keep working. The cheap fix this
week is also the thing that unblocks the strategic move.

Two caveats on option 8 worth carrying to the owner:

- **Publisher verification is close to mandatory.** Under the modern default policy
  (`microsoft-user-default-low`, "allow user consent for apps from verified publishers, for
  selected permissions"), an unverified multi-tenant app is excluded from the user-consent
  allowance, and risk-based step-up consent blocks users from consenting to newly registered
  unverified multi-tenant apps outright
  ([publisher verification](https://learn.microsoft.com/en-us/entra/identity-platform/publisher-verification-overview)).
  Agentic Delivery should be a verified publisher *before* anyone is asked to click consent.
- **ROPC has to go anyway, and its replacement has its own trap.** Microsoft: "Microsoft
  recommends you do not use the ROPC flow; it's incompatible with multifactor authentication"
  ([ROPC](https://learn.microsoft.com/en-us/entra/identity-platform/v2-oauth-ropc)). Device code
  flow is the obvious successor and is what `ropc-token-provider.ts`'s own doc comment already
  names — but Conditional Access has a dedicated **authentication-flows condition that blocks
  device code flow**, and Microsoft's own guidance recommends organisations block it by default
  ([block authentication flows](https://learn.microsoft.com/en-us/entra/identity/conditional-access/policy-block-authentication-flows)).
  We cannot see If's CA posture from outside. **Build against authorization-code flow**, which
  that condition does not target.

**Option 9 — push instead of polling — is available to us, contrary to what we assumed.**
A subscription on `/chats/{id}/messages` **is supported with delegated permissions**
(`Chat.Read` least-privileged;
[change notifications for Teams messages](https://learn.microsoft.com/en-us/graph/teams-changenotifications-chatmessage)),
maximum lifetime **4 320 minutes (3 days)** — the widely-repeated "60 minutes" figure is stale —
with `lifecycleNotificationUrl` mandatory whenever expiry is more than an hour out. Metering is
gone: "Starting August 25, 2025, the Teams APIs listed in this article are no longer metered, and
no billing configuration is required"
([Teams API licensing](https://learn.microsoft.com/en-us/graph/teams-licenses)), and single-chat
subscriptions were never in the metered set regardless. Delivery does **not** require a public
HTTPS endpoint: Azure Event Hubs delivery exists precisely for "when the receiver can't expose a
publicly available notification URL"
([Event Hubs delivery](https://learn.microsoft.com/en-us/graph/change-notifications-delivery-event-hubs)),
which suits a laptop/Ubuntu-server topology far better than opening a port.

One honest caveat found during the research and not resolved: an archived "Protected APIs in
Microsoft Teams" page listed "Create subscription for new chat messages" among application-only,
access-request-gated APIs. The live page 404s and the current permissions table contradicts it.
**Treat this as a cheap empirical test, not a doc question** — one subscription attempt answers it.

**Option 10, and why the bot is not simply "no".** A bot cannot post as a human: bot messages
carry the bot's own identity, and no supported mechanism sets `from.user` to an arbitrary person
(`Teamwork.Migrate.All` is bulk migration import, not live posting). Against our attribution rule
that disqualifies option 11 outright. But the *reading* half has no attribution requirement at
all. A bot installed in the chat receives every message by webhook — **zero polling, zero Graph
reads** — under documented, comparatively generous limits (50 rps per app per tenant; per-thread
send 7/s, 1 800/hour;
[bot rate limits](https://learn.microsoft.com/en-us/microsoftteams/platform/bots/how-to/rate-limit))
while every *post* continues to go out through delegated Graph as the named human. That hybrid is
the only design found that removes our entire read volume without touching attribution. Its cost
is the install: at an insurer, org-wide custom app upload is very likely off, in which case the
only route is "submit a custom app for admin approval" — a path Microsoft guarantees always exists
and cannot be disabled
([custom app policies](https://learn.microsoft.com/en-us/microsoftteams/teams-custom-app-policies-and-settings)),
but which is a formal ask of If's IT and highly visible in the Teams admin centre.

**On visibility, for the record.** Neither option 8 nor option 10 is covert, and neither should
be. A consent creates a service principal in Enterprise Applications plus a `Consent to
application` audit event; an app install appears in Teams admin centre → Manage apps. That is
consistent with how we already work — nothing here gets done without If knowing it was done.

## 4. Recommendation — a staged plan

### Stage 1 — this week, inside the library, no one else involved

1. **Record `x-ms-throttle-scope` on every 429.** Graph returns the header as
   `<Scope>/<Limit>/<ApplicationId>/<TenantId|UserId|ResourceId>` with scope values
   `Tenant_Application`, `Tenant` or `Application`. `GraphClient.noteThrottle` already inspects
   the response; capturing this header there, surfacing it through `GraphError`, and writing it to
   the daemon log turns section 2's careful inference into a measured fact in one throttled call.
   **Do this first** — every later decision gets cheaper if we know which of the four keys is
   closing.
2. **Take `/members` off the send path for good.** Harvest `(fromId, from)` from every polled
   message into the members cache at zero Graph cost; serve a stale roster when a refresh is
   throttled rather than failing the send; make TTL expiry mean "refresh opportunistically", never
   "the cache is empty". This is the direct fix for the 2026-09-04 incident *and* the prerequisite
   for stage 2's consent-free scope set.
3. **Poller supervision, landed 0.5.4.** PR #8 predated the 0.5.2 poller rewrite (quota yield,
   roster harvest, auth-health tracking) and no longer merged cleanly, so the work landed as a
   rebased branch carrying the same behaviour onto current `main` instead of a literal merge of
   that PR: the Retry-After-as-backoff-floor change, the escalating error lines, the health file
   (`poller-health.json`, `src/inbox.ts`), and the single-instance lock (`poller-lock.ts`,
   `poller.lock`, keyed per inbox path via `inboxPathFor(env)`). This is the evidence source this
   plan's later stages read from.
4. **Two knobs, immediately, while the above ships**: give the CTP and SRP daemons distinct
   `TEAMS_MCP_CLIENT_ID` values (they are identical today — both `d3590ed6…`), and raise the CTP
   poll interval to 180 s. Neither is a fix; both are free insurance against the sustained-83 %
   rule and the per-chat 1 rps ceiling.

### Stage 2 — next, on Agentic Delivery's own account, still nothing asked of If

5. **Register `Agentic Delivery Teams Assistant`** as a multi-tenant Entra app in the Agentic
   Delivery tenant, requesting **only user-consentable delegated scopes**: `Chat.ReadWrite`,
   `ChatMessage.Send`, `User.Read`. Deliberately **not** `ChatMember.Read` — that one permission
   is the difference between "the consultant consents for himself" and "we need If's Global
   Admin".
6. **Complete publisher verification** before anyone is asked to consent, or step-up consent will
   refuse the app regardless of how modest the scopes are.
7. **Ship an authorization-code token provider** alongside `RopcTokenProvider`. The
   `TokenProvider` seam already exists for exactly this swap. Auth-code rather than device-code,
   because CA can block device-code flow tenant-wide and we cannot see If's policy.
8. **Prototype change-notification ingestion** against the SRP tenant first — our own risk, our
   own tenant, no customer exposure — using Event Hubs delivery in Agentic Delivery's Azure
   subscription. This also settles the protected-API ambiguity empirically.

### Stage 3 — the one ask of the customer, made once, with the cost stated

9. **Ask for consent to the Agentic Delivery app in If's tenant.** Under a default consent policy
   the consultant can consent for himself and no admin action is needed; if user consent is
   restricted, the admin-consent-request workflow exists and the ask is a single approval of three
   low-impact delegated scopes — not a Global Admin operation, not a tenant-wide grant. **This is
   the change that actually moves us out of the shared bucket**, and everything in stage 1 is
   designed so we survive comfortably until it lands.
10. **Hold the bot-for-ingestion hybrid (option 10) in reserve.** Propose it only if, after stages
    1–2, reads are still being throttled — it is the strongest remaining lever but it costs a
    formal app-approval request to If's IT and a permanent, highly visible artifact in their
    Teams admin centre. Not worth spending that goodwill before the cheaper moves are measured.

**What we are explicitly not doing:** batching (no budget effect), one shared poller (the buckets
are already separate), delta query (closed to delegated auth), and a bot that posts (breaks
attribution).

## 5. Open questions for the owner

1. **Is a per-instance client id acceptable as an interim measure?** Both daemons currently
   present the Microsoft Office id. Splitting them is one env var, but it means two different
   Microsoft first-party identities appearing in If's sign-in logs against the same account. Is
   that better or worse from the "everything we do wears your name" standpoint?
2. **Who at If do we ask about consent policy and Conditional Access**, and do we ask before or
   after the app is registered and verified? Asking first is more honest; asking after means we
   arrive with something concrete rather than a hypothetical.
3. **Is publisher verification available to Agentic Delivery today** — does it hold a Microsoft AI
   Cloud Partner Program account and a non-`onmicrosoft.com` publisher domain? This gates stage 2
   entirely and has a lead time we do not control.
4. **How much inbox latency is acceptable?** Stage 1 proposes 180 s. Push notifications would take
   it back under 10 s, but if 180 s is fine, option 9's urgency drops considerably.
5. **Does the attribution requirement apply to reads?** It is written as a rule about writes. If
   reads may be attributed to an application rather than a person, option 10's hybrid becomes far
   more attractive and should be pulled forward.
6. **Is an Azure spend in the Agentic Delivery subscription approved** for an Event Hubs namespace
   plus a small always-on receiver? Modest, but it is a standing cost against the cost-conscious
   infrastructure rule.
