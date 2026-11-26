# Plan tab — server self-test checklist

The Apps Script has no test runner in this repo, so the server side is verified
manually. Run through these steps after deploying the updated
`SiteAcceptanceTracker.gs`.

Replace `<SCRIPT_URL>` with the deployed web-app URL, e.g.

```
https://script.google.com/macros/s/AKfycbxYb30QpZOkQSrRhx2hvuHPA8P3JMXxiqZAVGqb2b1UibXTUT-ZIvMO6cdJa2s7dhLy/exec
```

---

## 0. Deploy order

1. Paste the updated `SiteAcceptanceTracker.gs` into Apps Script → **Deploy → New version**.
2. **Before** opening the app, run the migration in the Apps Script editor:
   select `migratePlanRoster` from the function dropdown → **Run**.
   It prints a summary — check it before continuing.
3. Then open the site. The client will fetch the new roster document.

> The migration must run BEFORE the first client load, because the new client
> no longer writes to the legacy `planRosterEdits` map.

---

## 1. Roster document

```bash
curl -s "<SCRIPT_URL>?type=planRosterDoc"
```

**Expect:** `{"status":"ok","version":N,"accounts":[...]}`

- [ ] Every account has a non-empty `id`.
- [ ] Every `subcon` is UPPERCASE and one of:
      `ADW CAA CCSI INDOHR INNOVIS INTISEL MKU MPJ NIM PAB POCA TRITAMA YPIT ZMG`
- [ ] The `warsidi_*` family appears exactly ONCE, as `warsidi_poca_tl`.
- [ ] `TESTACC` and `FIXCHECK_1` are absent.
- [ ] `version` is a number ≥ 1.

## 2. Roster command — upsert (add a member)

```bash
curl -s -X POST "<SCRIPT_URL>" -H "Content-Type: text/plain;charset=utf-8" \
  -d '{"type":"planRosterCommand","op":"upsert","version":<VERSION>,"account":"selftest_user","subcon":"PAB","teamType":"WL","resourceRemark":"Shared"}'
```

**Expect:** `{"status":"ok","version":<VERSION+1>,"roster":[...],"warnings":[]}`

- [ ] `version` incremented by exactly 1.
- [ ] The new member is in `roster`.
- [ ] `warnings` is empty.

## 3. Roster command — stale version is rejected

Repeat step 2 but send `"version":1` (a stale value).

**Expect:** `{"status":"stale","version":<current>,"roster":[...],"message":"Roster changed on the server — please retry"}`

- [ ] `status` is `stale`.
- [ ] The roster is returned so the client can re-apply.

## 4. Roster command — rename keeps the id

```bash
curl -s -X POST "<SCRIPT_URL>" -H "Content-Type: text/plain;charset=utf-8" \
  -d '{"type":"planRosterCommand","op":"rename","version":<VERSION>,"id":"<ID>","newAccount":"selftest_user_renamed"}'
```

**Expect:** `status:ok`, the account name changed, **the `id` unchanged**.

- [ ] The `id` is identical to before the rename.

## 5. Roster command — remove

```bash
curl -s -X POST "<SCRIPT_URL>" -H "Content-Type: text/plain;charset=utf-8" \
  -d '{"type":"planRosterCommand","op":"remove","version":<VERSION>,"id":"<ID>"}'
```

**Expect:** `status:ok`, the member has a non-empty `removedAt`.

- [ ] The member is soft-removed (`removedAt` set), not deleted outright.

## 6. Plan rows carry rosterId

```bash
curl -s "<SCRIPT_URL>?type=plan&date=2026-11-26"
```

**Expect:** an array of row objects; each has a `rosterId` field.

- [ ] Rows written after the migration have a non-empty `rosterId`.
- [ ] Legacy rows (written before the migration) have `rosterId: ""` — these are
      matched by account name on the next save and then back-filled.

## 7. Plan save matches by rosterId

Open the Plan tab in the browser, type a value into a cell, press **Save**,
then reload the page.

- [ ] The value survives the reload.
- [ ] The status line shows a diff, e.g. `Saved (1 updated)`.
- [ ] Check the `plan` sheet: column L (`rosterId`) is populated for that row.

## 8. Plan save returns a diff

```bash
curl -s -X POST "<SCRIPT_URL>" -H "Content-Type: text/plain;charset=utf-8" \
  -d '{"type":"planSave","date":"2026-11-26","rows":[{"rosterId":"<ID>","account":"<ACCOUNT>","subcon":"PAB","teamType":"WL","resourceRemark":"Shared","duId":"","siteName":"","activityRemark":"","dailyPlanActivity":"","result":""}]}'
```

**Expect:** `{"status":"ok","written":N,"updated":N,"appended":N,"unchanged":N,"deleted":N,"warnings":[...]}`

- [ ] `updated` + `appended` + `unchanged` accounts for every submitted row.

## 9. Legacy row back-fill

Take a row that still has `rosterId: ""` from step 6, save the Plan tab, and
re-fetch step 6.

- [ ] That row now has a non-empty `rosterId`.

## 10. Subcont visibility

Log in with each subcont token and confirm every member of that team appears
in the Plan table, and no member of another team appears.

- [ ] `ADW` (2 members), `MPJ` (2 members), `MKU` (1 member) can now log in.
- [ ] `YPIT` can log in and sees an empty team (0 members) — expected.

---

## Cleanup

The self-test creates `selftest_user`. Remove it with:

```bash
curl -s -X POST "<SCRIPT_URL>" -H "Content-Type: text/plain;charset=utf-8" \
  -d '{"type":"planRosterCommand","op":"remove","version":<VERSION>,"id":"<ID>"}'
```

---

## Troubleshooting

| Symptom | Cause | Fix |
|---|---|---|
| `?type=planRosterDoc` returns 404-ish / `handleForward` error | The new Apps Script version is not deployed | Redeploy `SiteAcceptanceTracker.gs` |
| Migration prints nothing | `migratePlanRoster` was not selected/run | Select the function in the dropdown and press Run |
| Client shows "Could not load the roster" | The server is unreachable, or the client is talking to an old backend | Check the console for the exact request URL |
| Rows have empty `rosterId` | The migration has not back-filled the sheet yet | Run `migratePlanRoster`, then save the Plan tab once |
