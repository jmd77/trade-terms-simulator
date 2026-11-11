/*
 * Trade terms simulator — app.js (the whole app: shell, engine, Excel and all five screens)
 *
 * LAYOUT OF THIS FILE
 *   1. Core IIFE: example data (+ the monthly-row helpers), fmt, engine, state, ui helpers, Excel, routing + render, popover,
 *      then screen 1 Priorities, screen 2 Import, screen 3 Agreement (Contract analysis, Agreement today = the price waterfall
 *      and its optional By pack table, All customers = one chart), and TTS.start().
 *   2. Screen 4 Track & trace (own IIFE, registers TTS.screens.track; CSS classes prefixed tt-). Ninth pass (SPEC9): a calendar.
 *      Rows are read-only, no input in a row: KPI · calendar · Status · chevron, for one customer and for All customers alike.
 *      KPI cell (td.tt-kpi): .tt-l1.tt-name = k.text (the KPI in words) + .tt-l2 (13 px muted, flex): the contract point (All
 *      customers: a.tt-cust link to that customer's Track & trace, " · ", the contract point), then the .tt-ahead "Paid ahead of
 *      delivery" flag (alert icon 14 px, title: the two percentages) on every KPI of a line paid ahead. Calendar (td.tt-calcell):
 *      span.tt-nowmark (SPEC10 3: a light mark behind the demo month's cell in closed rows that have one, --cur = its index) + span.tt-cal[role=img][aria-label = the
 *      twelve months in words] > twelve i.tt-d.done|part|late|plan|none (colour only, title in words; state from monthState,
 *      SPEC9 2.3); its header is span.tt-cal.tt-calh > twelve .tt-ml labels (.now = the demo month, navy pill). Cell size from
 *      --cw / --cg on .tt-grid (34/4 px, 30/3 below 1280 px). div.tt-legend (three .tt-sw swatches) sits between the band and
 *      the table; band tile "Paid ahead of delivery" from E.paidAheadTotal. The chevron (td.tt-chcell > button.iconbtn.tt-chev[.open],
 *      data-action="trackDetails", data-href = the route to open or close, data-key track:<c>:<kpiId>:details, aria-label "Show
 *      the detail" / "Hide the detail").
 *      ONE panel is open at a time and it is the route: #track/<c>/<kpiId> (one customer) and #track/all/<c>/<kpiId> (All
 *      customers); TTS.route drops an id that is not a KPI of that customer (any more), the scope switch keeps the open KPI
 *      (#track/a/a6 <-> #track/all/a/a6), the panel is scrolled into view once when its route is reached (TTS.scrollTarget).
 *      #track/<c> and #track/<c>/none = every panel closed (closing goes to #track/<c>); #track/all opens none. All customers
 *      is sorted Not executed, Partially executed, Fully executed, then money descending. The whole closed KPI row
 *      (tr.tt-krow, same data-action and data-href as its chevron) opens its card; a click on a link or button in it does not.
 *      SPEC10: the open KPI is ONE row instead, tr.detrow.tt-detrow[data-kpi] > td[colspan] > div.tt-card (border, shadow,
 *      clipped corners) = div.tt-cardhead (grey; same data-action and data-href, closes the card unless the click is on a link or
 *      button in it; a grid on the table's column widths: .tt-kpi, an empty calendar slot (no month cells), .tt-hst status chip,
 *      .tt-chcell chevron) + .ca-det.tt-det (white; thin dividers between the blocks), three section.ca-blk:
 *        .ca-b1 What was agreed, plain text: p.tt-agreed (sentence), p.tt-money ("SEK 2.5 m depends on it."), p.tt-src
 *          ("Tracked through: … · Owner: …", empty parts left out), a.tt-edlink -> #agreement/<c>/analysis/<lineId>.
 *        .ca-b2 Month by month: .tt-mm > .tt-mhead (title + button.tt-edit "Edit numbers" | "Done", data-action="trackEdit",
 *          data-key track:<c>:<kpiId>:edit; which KPI is in edit mode is view state only) + .tt-strip (12-column grid: .tt-ml
 *          month labels, .now = the demo month) + p.tt-msum (the summary line). Strip: one .tt-mc per month — button (ended or
 *          current month: data-action="trackMonth", data-key track:<c>:<kpiId>:m:<m>, aria-pressed = something recorded; a
 *          click records plan[m] (the target level for a %), a second click clears it) or span (future month, or .none = nothing
 *          planned "–"); states .done (green + .tt-tk tick; reached, a % at the month's target or more) .part (yellow, "1/2") .late
 *          (ended, planned, nothing recorded) .plan (grey); .sm = a longer number in 13 px. Edit mode: .tt-strip.edit with .tt-rl
 *          row labels (Planned / Executed; Target / Measured for a %) and input.tt-mi (data-change="month", data-key
 *          track:<c>:<kpiId>:plan:<m> / done:<m>).
 *        .ca-b3 Proof and payment: .tt-pp (label grid, .tt-f > .tt-fl + .tt-fc): select proof (E.PROOFS), select the line's
 *          paidHow, input.tt-paidin "Paid so far" (data-change="line" data-field="paid" data-num, data-key
 *          track:<c>:<kpiId>:paid) + "of 2.5 (SEK m)"; then .tt-sws > four button.tt-switch[role=switch][aria-checked]
 *          (data-action="trackSwitch", data-field tracked | checked | due | paidFull, data-key track:<c>:<kpiId>:<field>).
 *      Laptops (< 1500 px): What was agreed then Month by month in the left column, Proof and payment in the right one.
 *      Cards take their content height (no inner scroll); below 1400 px the script card takes its own row;
 *      the opportunity total (.pl-oppsum) shows only when card 2 stays no taller than the cards beside it (fitOppSum);
 *      All customers: one .pl-headband headline above the cards.
 *      Fifth pass (SPEC5 8): the key figures come first, from E.keyFigures: nav.pl-kfband with five a.pl-kf groups (.pl-kf-terms
 *      -> #agreement/<c>/waterfall, -cls and -cps -> #agreement/<c>, -kpis and -ahead -> #track/<c>; All customers: #agreement/all
 *      and #track/all; data-key plan-kf:<group>): Trade terms (SEK, % of gross revenue, third net), Classification (.pl-kfbar
 *      + .pl-kfcls legend), Counterparts, KPIs (.pl-kfchips), Paid ahead of delivery (.pl-kfahead / .pl-kfnone). Five in a row
 *      from 1500 px, three then two below. All customers: the headline, then the same band for the total. Print: the band is
 *      hidden and .pl-kfprint (the same figures as one line under the print title) is shown.
 *      Script edits are kept per customer in customer.scriptEdits[key], keys: opening, propose, ask, give,
 *      obj1q, obj1a … obj4q, obj4a ("ask" and "give" are the two columns of "What we ask and what we give").
 *   Screens 4 and 5 only use the core through TTS.* (no shared private variables).
 *
 * PUBLIC API (everything lives on window.TTS)
 *   TTS.state                      the live state (SPEC section 3 + SPEC5 1 + SPEC5b A: version 5, see DATA MODEL below)
 *   TTS.screens[name]              { head(scope, customer, route) -> {title, subtitle, actions},
 *                                    one(customer, route) -> html, all(route) -> html, afterRender?(root, route) }
 *   TTS.actions[name](el, event)   delegated clicks on [data-action="name"]
 *   TTS.changes[name](el, event)   delegated changes on [data-change="name"]
 *   TTS.commit()                   save to localStorage + render
 *   TTS.render()                   rebuild the page from state (keeps focus by data-key, keeps scroll)
 *   TTS.go(screen, scope, option)  navigate (hash routing #screen/scope[/option], option may hold a "/", e.g. "waterfall/proposed")
 *   TTS.route()                    current route { screen, scope, option, sub, view, packs, fixed }: #screen/scope/option/sub;
 *                                  view = 'analysis' | 'waterfall' on step 3 of one customer, '' elsewhere; packs = true on
 *                                  #agreement/<c>/waterfall/today|proposed/packs (then sub is '' or 'proposed')
 *   TTS.customer(id), TTS.scopeLines(scope), TTS.uid(prefix)
 *   TTS.hashFor(screen, scope, option) -> "#screen/scope[/option]" · TTS.listJoin(["a","b","c"]) -> "a, b and c"
 *   TTS.dayNum("2027-06-15") -> whole days since 1970-01-01 (UTC), for day-of-year arithmetic
 *   TTS.load(), TTS.save(), TTS.resetToExample(), TTS.exampleState()   (saved under localStorage "tts.state.v5", state.version 5;
 *                                  older keys are ignored, load removes "tts.state.v4"; load fills every missing field with
 *                                  its default)
 *   TTS.applyAction(line, action)  sets line.action and applies its effect once (SPEC5 3.2): Keep -> proposed amount, type and
 *                                  priority = today's · Make conditional, Strengthen the counterpart -> proposed type conditional ·
 *                                  Fold into the list price, Remove -> proposed amount 0 · Invest more, Reduce, Move to a priority ->
 *                                  nothing automatic. Returns the line. (The caller commits.)
 *   TTS.recalc(customer?)          SPEC5b A: THE ONE PLACE that recalculates the amounts of the variable lines, amount = rate ÷ 100 ×
 *                                  base, one decimal: on-invoice variable lines first (base: gross revenue = customer.listSales),
 *                                  then the first net (listSales − every on-invoice amount), then every other variable line (base:
 *                                  the first net). A variable line with rate null takes its rate from its amount (two decimals).
 *                                  With packs: promoSpend = the sum of the packs' promo (one decimal). No argument = every customer.
 *                                  Returns the customer (the state without argument). Called by load (normCustomer), the import
 *                                  parser and TTS.commit() — so after any edit just call TTS.commit(); everything else keeps
 *                                  reading line.amount.
 *   TTS.setBasis(c, line, basis)   'Fixed' | 'Variable' (SPEC5b A): to Variable the rate becomes E.rateOf (the current rate, two
 *                                  decimals, so the amount stays); to Fixed the amount stays and rate becomes null. Caller commits.
 *   TTS.setPromoSpend(c, value)    the customer's promotional spend (≥ 0, one decimal). With packs, the typed total rescales each
 *                                  pack's promo in proportion (by share when they are all 0), one decimal, the rounding rest on the
 *                                  largest, so their sum = the total. Caller commits.
 *   TTS.applyImport(parseResult)   adds or replaces the customer of a parse result (same name, case-insensitive), returns it
 *   TTS.exportStyled(scope)        NOT defined here: export.js (Part 3) adds it; the top bar's "Export to Excel" calls it when it
 *                                  exists, otherwise the old two-sheet SheetJS export (TTS.excel.exportWorkbook)
 *   TTS.details                    view state of step 3: { "<customerId>:<lineId>": true } for open Details rows (never saved)
 *   TTS.scrollTarget               a CSS selector a screen may set during render: that element is scrolled into view (block
 *                                  'nearest') once, after the render restored the scroll position (used by Track & trace)
 *   TTS.plan     generate(customer), texts(customer), plainText(customer)  (screen 5: script text)
 *
 * ROUTES OF STEP 3 (SPEC5 4)
 *   #agreement/<c> or #agreement/<c>/analysis       Contract analysis (default view of one customer)
 *   #agreement/<c>/analysis/<lineId>                the same with that line's Details open (applied once when the route is
 *                                                   reached, scrolled into view; closing that Details returns the hash to #agreement/<c>)
 *   #agreement/<c>/waterfall[/proposed]             Agreement today / proposed (SPEC5c 3): the band and the price waterfall, the
 *                                                   picture of the Contract analysis table (same rows: E.waterfall), read-only;
 *                                                   clicking a line goes to #agreement/<c>/analysis/<lineId>
 *   #agreement/<c>/waterfall/today/packs            By pack (SPEC5b G), only for a customer with pack groups (else the route
 *   #agreement/<c>/waterfall/proposed/packs         falls back to the waterfall): one table from E.packWaterfall, the pack
 *                                                   groups' share and promotional spend typed there. Control "All packs | By pack"
 *                                                   (data-keys packs-all, packs-by) in the first header cell of the block.
 *   #agreement/all                                  All customers (SPEC5 6.2): only the chart (no Today | Proposed; any option is
 *                                                   dropped, e.g. the old #agreement/all/proposed -> #agreement/all). The view
 *                                                   switch has a third option "Compare customers" (data-key view-all) = this chart;
 *                                                   on the chart the other two lead to the customer selected last (ui.lastCustomerId,
 *                                                   else the first). Each customer's pair is a link to #agreement/<c> (Contract
 *                                                   analysis). On step 3 the head's controls wrap to their own row under the
 *                                                   subtitle when they do not fit beside the one-line title (CSS).
 *   The old #agreement/<c>/proposed becomes #agreement/<c>/waterfall/proposed. Coming from another screen, step 3 opens on Contract
 *   analysis. Next: Contract analysis -> "Next: Agreement today" -> "Next: Track & trace". View switch data-keys: view-analysis,
 *   view-waterfall; Today | Proposed: mode-today, mode-proposed (they keep /packs).
 *
 * ENGINE — TTS.engine (E). mode is 'today' | 'proposed'. Amounts in SEK m; shares and pct are fractions (0.124 = 12.4%).
 *   amount(line,mode), type(line,mode), priorityId(line,mode), priority(id), hasNumber(line) (counterpart, target, date),
 *   missing(line) -> ['counterpart','target','date','owner'] still lacking (missing(kpi, 'name') for an extra KPI: no owner check)
 *   classOf(line, mode) -> 'unconditional' | 'partial' | 'conditional'   (SPEC5 3.3: conditional needs counterpart, target number,
 *     date AND owner; otherwise a conditional-type line is 'partial'). CLASS_LABEL = { unconditional: 'Unconditional',
 *     partial: 'Partially conditional', conditional: 'Conditional' }. lineStatus is an alias of classOf (no more 'nonumber').
 *   smart(line, mode) -> { specific, measurable, owner, timebound, score }  (all false, score 0 when the type is unconditional)
 *   aligned(line, mode) -> true when the line has a priority in that view
 *   suggestPriority(line) -> priority id or null: only for a line conditional TODAY, with no priority today and
 *     priorityDecided === false (word match on name + counterpart + bucket hint, SPEC5 3.3). Counts nowhere until confirmed.
 *   ACTIONS (the eight action names, in menu order) · PROOFS (the seven proofs) · MONTHS ['Jan' … 'Dec']
 *   summary(lines, mode) -> { total, unconditional, partial, conditional, linked, shareUnconditional, sharePartial,
 *     shareConditional, shareLinked } (+ nonumber / shareNoNumber kept as aliases of partial / sharePartial)
 *   linkedShare(lines, mode) -> { amount, total, share }   (spend with a priority in that view)
 *   sales(customer | [customers], mode) -> { gross, onInvoice, firstNet, offInvoice, secondNet, promo, thirdNet, tradeTerms }
 *     (SPEC5b B; SEK m; several customers = the sums): gross = gross revenue at list price (today customer.listSales; proposed
 *     listSales minus today's amount of every line whose action is 'Fold into the list price') · onInvoice = the view's amounts
 *     of the lines paid 'On invoice' · firstNet = gross − onInvoice (net invoice sales) · offInvoice = the view's amounts of all
 *     other lines (Off invoice, Fixed amount, blank) · secondNet = firstNet − offInvoice · promo = customer.promoSpend (same in
 *     both views) · thirdNet = secondNet − promo (net-net) · tradeTerms = onInvoice + offInvoice (total of the contract points).
 *     The old names listSales / netInvoice / netNet / totalDiscount and E.pctOfNet are gone.
 *   pctOfGross(amount, customer | [customers], mode) -> amount ÷ gross of that view, a fraction (0 when gross <= 0). "% of gross
 *     revenue": every waterfall row, and the totals in bands and key figures.
 *   rateBase(line) -> 'gross' (paid On invoice) | 'firstNet' (every other line): the key in E.sales() of the line's contract base
 *   rateOf(line, customer) -> the line's rate in % of its base, a PERCENT NUMBER (1.72 = 1.72%), two decimals: the stored
 *     line.rate of a Variable line; for a Fixed line (or a Variable one with rate null) today's amount ÷ today's base × 100.
 *     null when the base is not above 0. Show it with fmt.rate().
 *   waterfall(customer, mode) -> THE ONE PRICE STRUCTURE (SPEC5c 1): the Contract analysis table (with mode 'today') and the
 *     Agreement today picture (either mode) are both drawn from it: same rows, same order, same numbers. Rows in this order:
 *       1. { kind: 'anchor', key: 'gross', label: 'Gross revenue at list price', hint: null }
 *       2. one { kind: 'line', block: 'on' } per contract point paid 'On invoice', in the customer's line order
 *       3. { kind: 'anchor', key: 'firstNet', label: 'First net (net invoice sales)', hint: 'After the on-invoice contract points' }
 *       4. one { kind: 'line', block: 'off' } per other contract point (Off invoice, Fixed amount, blank), in line order
 *       5. { kind: 'anchor', key: 'secondNet', label: 'Second net', hint: 'After the off-invoice contract points: rebates, bonuses
 *          and fixed allowances' }
 *       6. { kind: 'promo', key: 'promo', label: 'Promotional spend (estimate)', hint: 'Price discounts in promotions. Not a
 *          contract point.' }
 *       7. { kind: 'anchor', key: 'thirdNet', label: 'Third net (net-net)', hint: 'After promotional spend' }
 *     EVERY row has exactly these keys, in this order: { kind, key, block, label, hint, amount, pct, lineId, bucket, bucketKey,
 *     cls, basis, rate }. amount = SEK m in that view, positive (a deduction: show it negative); anchors = E.sales(c, mode)
 *     gross / firstNet / secondNet / thirdNet, promo = customer.promoSpend; pct = amount ÷ gross revenue of that view (fraction;
 *     gross = 1). Line rows: key null, block 'on' | 'off', label = line.name, hint null, lineId, bucket (E.bucketOf(line, mode)
 *     name), bucketKey (its E.BUCKETS key, colour class .bk-<key>), cls (E.classOf(line, mode)), basis ('Fixed' | 'Variable'),
 *     rate (E.rateOf(line, c): today's rate in % of its base, a percent number, or null). Non-line rows: block, lineId, bucket,
 *     bucketKey, cls, basis, rate = null. EVERY contract point has a row in both views, also at 0 (folded into the list price,
 *     removed): the 'today' and 'proposed' line rows are the same lines in the same order. No group, bucket or total rows any
 *     more (Total trade terms = E.sales().tradeTerms). Changing a line's "Paid how" moves it to the other block.
 *   packWaterfall(customer, mode) -> null when the customer has no packs; else [All packs, then one column per pack group in
 *     customer.packs order], each { name, gross, onInvoice, offInvoice, promo, firstNet, secondNet, thirdNet, pct: { gross,
 *     onInvoice, firstNet, offInvoice, secondNet, promo, thirdNet } } (SEK m; pct = fractions of THAT column's gross). A pack
 *     group's gross, onInvoice and offInvoice = the view's totals × share ÷ 100; its promo = its own promo. All packs = E.sales.
 *   keyFigures(customer | [customers] | nothing = all) -> { total, pctOfGross, sales, classes: { unconditional, partial,
 *     conditional } (each { amount, share }), counterparts: { lines, smart, aligned, smartNotAlignedAmount }, kpis: { n, executed,
 *     partly, notexecuted, atStake }, paidAhead, linkedToday, linkedProposed } — all today except linkedProposed; pctOfGross =
 *     total ÷ gross revenue (fraction); sales = E.sales(today) of that customer or the sums; linked* are linkedShare() objects
 *     { amount, total, share }; counterparts.lines = conditional-type lines today, smart = score 4.
 *   spendByPriority, kpis, allKpis, moneyAtStake, moneyByStatus, growth (the first net today against revenueLY),
 *   opportunitySize, suggestedMustWin, mustWin, opportunities (kinds: unconditional, partial, nopriority, growth, gap),
 *   elapsed, statusLabel(status), isPercent(unit), paidAhead(customer), paidAheadTotal(customers),
 *   BUCKETS, bucketOf(line, mode) ('Unconditional' unless classOf is 'conditional'), spendByBucket(lines, mode)
 *   Monthly rows (SPEC5 3.4): periodMonths(start, end) -> 12 booleans (months the period touches) · defaultPlan(kpiLike) -> 12
 *     numbers (a count spread evenly, remainder to the first months; a percentage repeated) · monthsEnded() -> months fully
 *     ended on the demo date (5 on 15 June) · currentMonth() -> index of the demo month (5; -1 before the year) ·
 *     doneFromActual(kpiLike, actual) -> 12 values · actualOf(kpiLike) -> count: sum of recorded done; percentage: latest
 *     recorded month; null when nothing is recorded · expectedOf(kpiLike) -> sum of plan over the ended months (null for a %).
 *   each KPI from kpis(customer) (lines classOf 'conditional' today: the line's own KPI, id = line id, then each complete
 *     extra KPI): { id, lineId, lineName, customerId, customerName, name, text, target, unit, start, end, source, owner, money,
 *     plan, done (copies), proof, tracked, checked, due, paidFull, actual, expected (null for a %), status, kind, delivered,
 *     isExtra, elapsed, paid, paidShare, lineDelivered, ahead, excess }. text (SPEC9 2.1) = kpiFull trimmed, or composed:
 *     count "10 leaflet weeks…" (target + name, first letter lower-cased), percentage "Name: 85%". Status (SPEC9 2.2):
 *     notexecuted "Not executed" (nothing recorded, or it adds up to 0); executed "Fully executed" (a count with actual >=
 *     expected, a percentage with its latest measured month >= target); partly "Partially executed" (anything in between).
 *
 * DATA MODEL (SPEC5 1)
 *   customer: { id, name, channel, example, sourceNote, listSales (gross revenue at list price this year, SEK m), revenueLY (net
 *     invoice sales last year), investmentLY, promoSpend, packs?, mustWinId, scriptEdits, lines }   — revenueTY is gone
 *     promoSpend (SPEC5b A): promotional spend this year, SEK m, the user's estimate (0 by default); not a contract point
 *     packs (optional; absent = no pack groups; imported customers have none): [{ name, share (% of gross revenue; the shares
 *       should add up to 100, nothing forces it), promo (SEK m) }]; with packs, promoSpend = the sum of promo (TTS.recalc)
 *   line: { id, name, amount, paidHow, bucket, type, priorityId, propAmount, propType, propPriorityId, counterpart, target, unit,
 *     start, end, source, owner, paid, clause, justification, action, why, priorityDecided, basis, rate, plan, done, proof,
 *     tracked, checked, due, paidFull, kpis }
 *     basis (SPEC5b A): 'Fixed' (the amount is typed, the rate is calculated) | 'Variable' (the rate is typed, amount follows
 *       in TTS.recalc) · rate: % of the base, two decimals (1.72), used when Variable; null on a Fixed line. propAmount stays
 *       typed in SEK m for every line.
 *     paidHow: 'On invoice' | 'Off invoice' | 'Fixed amount' | '' · paid: paid to date, SEK m, number | null
 *     bucket: 'Listings' | 'Shelf space' | 'In-store marketing' | 'Promotions' | 'Efficiency' | '' (one field for both views,
 *       counted only while the line is conditional in that view)
 *     clause, justification ('' = none), why: text · action: one of E.ACTIONS · priorityDecided: false while the priority is
 *       still open (then the tool may suggest one); picking anything in the priority select or confirming sets it true
 *     THE MONTHLY FIELDS ARE STORED FLAT on the line (its own KPI) and the same way on each extra KPI:
 *       plan: 12 numbers Jan–Dec · done: 12 x (number | null) · proof: one of E.PROOFS · tracked, checked, due, paidFull: booleans.
 *       There is no stored "actual" any more. A new target, unit or period gives that KPI defaultPlan() (done stays).
 *       kpiFull (SPEC9 2.1): the KPI in full words, '' by default (then composed); cleared when the counterpart (a line), the
 *       name (an extra KPI), the target or the unit changes. Excel import layout: optional last column "KPI in full".
 *   extra KPI (line.kpis[]): { id, name, target, unit, start, end, source, owner, plan, done, proof, tracked, checked, due, paidFull,
 *     kpiFull }
 *
 * BUILT-IN CHANGE HANDLERS
 *   data-change="line" data-cust="<customerId>" data-line="<lineId>" [data-kpi="<kpiId>"] data-field="<field>" [data-num="1"]
 *     -> sets the field on the line (or extra KPI) and commits. Special fields: priorityId (also sets priorityDecided),
 *     action (calls TTS.applyAction), proof (must be in E.PROOFS), tracked / checked / due / paidFull (value "yes" or "true"
 *     -> true, anything else false), target / unit / start / end (also resets the plan to the default), basis ('Fixed' |
 *     'Variable', through TTS.setBasis), rate (with data-num="1": a trailing "%" and a comma accepted, two decimals, blank = 0).
 *     Every commit runs TTS.recalc, so amounts, bands and shares follow at once; focus is kept by data-key.
 *   data-change="customer" data-cust data-field="listSales" | "promoSpend" | "revenueLY" | "investmentLY"
 *     -> a number (comma accepted; bad text refused with a toast; blank = 0); promoSpend goes through TTS.setPromoSpend. Commits.
 *     Used by the Contract analysis anchor inputs "Gross revenue at list price" and "Promotional spend (estimate)" (data-keys
 *     cust:<c>:listSales, cust:<c>:promoSpend). The Agreement today picture is read-only.
 *   data-change="pack" data-cust data-pack="<index in customer.packs>" data-field="share" | "promo"
 *     -> a number ≥ 0 (comma and a trailing "%" accepted; bad text refused; blank = 0; promo one decimal); promoSpend becomes
 *     the sum in the commit. Shares are not forced to 100: show the note when they do not add up.
 *   data-change="month" data-cust data-line [data-kpi] data-row="plan" | "done" data-month="0".."11"
 *     -> one monthly cell (comma accepted; bad text refused with a toast; blank = 0 in plan, null in done) and commits.
 *   Every editable control needs a unique, stable data-key (focus is restored by it after each render).
 *   TTS.fmt      m, money ("SEK 12.0 m"), pct, pct1, pct1s (signed), delta, num, parse, target, period, date, isoWeek,
 *                rate (a percent number: 5.3 -> "5.3%", 1.72 -> "1.72%", 1 -> "1.0%"; '' for null)
 *   TTS.ui       esc, icon (adds 'cross'), chip, statusChip, lineChip(classOf value), statBand(items, rightHtml),
 *                splitBar(summary, {legend}), toast, confirm({title,text,okLabel}) -> Promise<boolean>, priorityOptions(selectedId),
 *                emptyState({icon,title,text,button:{label,href|action}}), bottomRow(leftHtml),
 *                bucketChart(lines) (the two bars Today/Proposed on one scale + the legend), bucketRowBar(lines, mode)
 *                statBand items: { value, pre, unit, label, sub, good, cls } (cls = extra class on that .stat)
 *                bottomRow(left): the row under a screen's content, left content + the Next button (keeps the scope,
 *                none on the last step). render() appends bottomRow('') to any screen whose html has no .bottom-row.
 *   TTS.excel    templateWorkbook, customerWorkbook, exportWorkbook, exportName, parseCustomerWorkbook (XLSX first arg)
 *                Customer sheet rows: Customer, Channel, Revenue last year (SEK m), Sales at list price (SEK m), Investment last
 *                year (SEK m), Promotional spend (SEK m) (blank = 0); an older file with "Revenue this year (SEK m)" imports list
 *                price sales = that + the on-invoice lines. Lines sheet: "Fixed or variable" and "Rate (%)" right after Bucket
 *                (blank = Fixed; anything else: warning, Fixed; a Variable line with a rate gets its amount calculated, the
 *                amount cell may be blank; without a rate the rate comes from the amount; the rate is ignored on a Fixed line).
 *                A blank proposed amount = today's (calculated) amount; the action is derived after the recalculation.
 *                Pack groups are never imported. More KPIs sheet unchanged ("Actual to date" written from actualOf, read into
 *                done). Files without the new columns or row import as before, with no message.
 *
 * SHARED CSS CLASSES (styles.css)
 *   .card (panel card) + .label (green section label) · .label.muted (table-header style label)
 *   .chip .grey | .green | .yellow | .outline
 *   .btn .primary | .secondary | .ghost | .onnavy · .iconbtn (32 px round icon button) · .linkbtn (inline ghost link)
 *   table.grid (data table) · th.num / td.num (right-aligned) · td.edit (cell holding an inline input)
 *   tr.flag (yellow "Partially conditional" row) · tr.total · tr.sub (44 px sub-row)
 *   .cell (inline input / select / button looking like text) · .cell.num · .cell.strong · .cell.cta (yellow CTA placeholder)
 *   .cell.wrapedit (one-row <textarea> that wraps and grows; Enter commits, a line break is never stored)
 *   label.cell.ovsel > .ovtext + select (visible wrapping label, invisible native select on top; .ovcaret after the last word)
 *   .tgt (target number .tnum + unit .tunit) · td.dashcell · .detbtn (Details toggle, .open) · .propamt > input + .chg.up/.down
 *   .bottom-row (every screen: 24 px under the content, ghost button or note left, a.nextbtn right; hidden in print)
 *   .br-narrow (a <br> shown only below 1500 px) · .stat .t > .tsep
 *   .stats + .stat (.v value, .u unit, .l label, .t "today" line, .stat.good) · .split (split bar + .legend)
 *   .segctl (segmented control, children .on) · .empty (empty state) · .muted · .hl (green key phrase)
 *   .pillfield (40 px pill input/select with 1.5 px navy border) · .dot (10 px legend dot: .c .n .u .navy .blue .green .grey)
 *   Buckets: .bk-<key> (bucket colour, on bar segments and .dot) · .bk-bar (stacked bar, segments <i>) · .stats.bkband[.prop]
 *   > .bk-chart > .bk-bars (.bk-lab + .bk-bar, twice) + .bk-legend (grid, 4 cells per item: .bk-name, .bk-num, .bk-to, .bk-num
 *   [.better | .zero])
 *   Contract analysis (prefix ca-): .stats.ca-band (> .ca-cls > .stat.ca-mini, three small stats) · table.grid.agreement.ca
 *   (col.ca-c-fv/-amt/-cls/-bkt/-pri/-str/-act/-prop/-det; columns Contract point | Fixed or variable | SEK m · % of gross |
 *   Classification | Bucket | Priority | Strength | Proposed action | Proposed SEK m | Details), body rows in the order of
 *   E.waterfall(c, 'today'): tr.ca-anchor.ca-k-<key> (gross, firstNet, secondNet, thirdNet) and tr.ca-promo, each > td.ca-alab
 *   (colspan 2: .ca-an label + .ca-ah hint) · td.ca-amtcell > .ca-amt > (input.cell.num.ca-ain [gross, promo] | .ca-av) + .ca-pg
 *   · td (colspan 5, empty) · td.ca-aprop > .ca-apv (the proposal's figure) [+ .ca-fold "−0.6 folded in"] · td (empty);
 *   tr.ca-row.ca-on|.ca-off[.flag][data-line] (a contract point) · .ca-pt (name editor + .ca-sum summary: counterpart · target ·
 *   period, plus a yellow "no owner" only when the owner is missing; .ca-miss yellow missing piece) · td.ca-fvcell > .ca-fv >
 *   label.ovsel.ca-basis (Fixed | Variable) + input.cell.ca-rin (the rate, Variable only, title = its base) · td.ca-amtcell >
 *   .ca-amt > (input.cell.num [Fixed] | .ca-amtro [Variable: calculated text]) + .ca-pg (% of gross revenue) · .ca-basisnote
 *   (the sentence under "Paid how" in Details) · .qm (.qm-rate, .qm-str: the "?" with a title) · .ca-str > .ca-pips >
 *   .ca-pip[.on] + .ca-al[.on] > svg.ca-flag (strength) · label.ovsel.sug + button.ca-ok (suggestion + confirm) ·
 *   tr.detrow.ca-detrow[data-line] > .ca-det > section.ca-blk .ca-b1/.ca-b2/.ca-b3 (each > h3.label; .ca-f = label .ca-l +
 *   control .ca-c; .ca-kpis > .ca-kpi > .ca-kpih + .ca-kpig; .ca-addkpi; .ca-smart > .ca-cks > .ca-ck[.on] + .ca-aln; .ca-nocp;
 *   .ca-acts)
 *   Agreement today (prefix wf-): .stats.bkband.wf-band[.prop] (three stats, .wf-l2 / .wf-today extra label lines, then
 *   bucketChart) · .wf-block > table.wf (col.wf-c-lab/-amt/-pct/-bar; th.wf-th0 holds the All packs | By pack .segctl.small) ·
 *   tr.wf-anchor.wf-k-<key> (.wf-an label, .wf-ah hint) · tr.wf-line.click[.wf-zero] (data-action go to the analysis/<lineId>
 *   route; .wf-ln > .dot.bk-<key> + a.wf-n + .wf-tag + .chip.wf-chip) · tr.wf-promo (.dot.wf-pdot) · td.wf-amt, td.wf-pct ·
 *   td.wf-barcell > .wf-track > .wf-guide (faint line at each anchor value) + .wf-bar (anchor, from the left edge, .wf-brk break
 *   mark) | .wf-seg.bk-<key> | .wf-seg.wf-pseg (floating steps) · .wf-cap (caption in the bottom row)
 *   By pack (prefix pk-): table.grid.pk > tr.pk-share, tr.pk-net | tr.pk-ded (+ .pk-<key>) > td.pk-lab + td > .pk-cell >
 *   (.pk-amt | input.cell.num.wf-pin) + .pk-pct [+ .pk-bar > i on the third net] · .pk-sum[.off] · .pk-note (shares ≠ 100%)
 *   All customers (prefix ac-): .ac-chart > .ac-pairs > (a.ac-pair | div.ac-pair.ac-all) > .ac-hl (.ac-hll + .ac-hlv) + .ac-cols >
 *   .ac-col (role img) > i.ac-seg.bk-<key>[.dk] (> b share) + .ac-cl + .ac-name + .ac-ch + .ac-tot; .ac-sep; .ac-legend
 */
var TTS = (function (global) {
  'use strict';

  var HAS_DOM = typeof document !== 'undefined' && typeof window !== 'undefined';
  var STORE_KEY = 'tts.state.v5';
  var TTS = { screens: {}, actions: {}, changes: {}, state: null, lastImport: null, pop: null };

  var SCREENS = [
    { id: 'priorities', label: 'Priorities' },
    { id: 'import', label: 'Import' },
    { id: 'agreement', label: 'Agreement' },
    { id: 'track', label: 'Track & trace' },
    { id: 'plan', label: 'Negotiation plan' }
  ];

  var uidCount = 0;
  function uid(prefix) { uidCount += 1; return (prefix || 'x') + Date.now().toString(36) + uidCount.toString(36); }
  function blank(v) { return v === null || v === undefined || String(v).trim() === ''; }
  function finite(v) { return typeof v === 'number' && isFinite(v); }
  function clone(o) { return JSON.parse(JSON.stringify(o)); }
  function listJoin(arr) {
    if (arr.length <= 1) return arr.join('');
    return arr.slice(0, -1).join(', ') + ' and ' + arr[arr.length - 1];
  }

  // ---------------------------------------------------------------- example data
  // Three invented customers (SPEC5 2): A and B (Home), C (Away from Home). Lines, amounts, buckets, priorities, counterparts,
  // targets, dates, owners, data sources, paid how, paid to date and proposals as in the fourth pass, except B's Damage and returns
  // allowance (folded into the list price, proposed 0.0) and A's Marketing support counterpart (SPEC5 2.3).
  var PAID_HOW = ['On invoice', 'Off invoice', 'Fixed amount'];
  // the five buckets a line's money can be for (SPEC4 1); "Unconditional" and "No bucket" are never chosen by hand
  var BUCKETS = ['Listings', 'Shelf space', 'In-store marketing', 'Promotions', 'Efficiency'];
  // the proposed action of a contract point (SPEC5 1) and the proof of execution of a KPI
  var ACTIONS = ['Keep', 'Invest more', 'Make conditional', 'Strengthen the counterpart', 'Move to a priority', 'Reduce', 'Fold into the list price', 'Remove'];
  var PROOFS = ['Store visit photo', 'Leaflet copy', 'Shelf audit', 'Customer report', 'Sell-out data', 'Listing file', 'None'];
  var ANSWERS = ['tracked', 'checked', 'due', 'paidFull'];   // KPI tracked? Checked? Payment already due? Paid in full?
  var BASES = ['Fixed', 'Variable'];                           // SPEC5b A: line.basis
  function r2(n) { return Math.round((+n || 0) * 100 + (n < 0 ? -1e-9 : 1e-9)) / 100; }
  var MON = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

  // ---- monthly rows (SPEC5 3.4), pure functions of (start, end, target, unit, year, demo date)
  function dnum(iso) { var p = String(iso).split('-'); return Date.UTC(+p[0], +p[1] - 1, +p[2]) / 864e5; }
  function rep(v) { var a = []; for (var i = 0; i < 12; i++) a.push(v); return a; }
  function isPct(unit) { return String(unit || '').trim().charAt(0) === '%'; }
  // which months of the year the period touches (blank start = 1 Jan; no end = no month)
  function monthsIn(start, end, y) {
    var out = rep(false);
    if (!end) return out;
    var s = dnum(start || (y + '-01-01')), e = dnum(end);
    for (var m = 0; m < 12; m++) out[m] = s <= Date.UTC(y, m + 1, 1) / 864e5 - 1 && e >= Date.UTC(y, m, 1) / 864e5;
    return out;
  }
  // a count is spread evenly over the months of its period (whole numbers, the remainder to the first months);
  // a percentage target is repeated in each month of the period
  function planFor(o, y) {
    var plan = rep(0), inP = monthsIn(o.start, o.end, y), months = [];
    inP.forEach(function (x, m) { if (x) months.push(m); });
    if (!finite(o.target) || !months.length) return plan;
    if (isPct(o.unit)) { months.forEach(function (m) { plan[m] = o.target; }); return plan; }
    var total = Math.max(0, Math.round(o.target)), base = Math.floor(total / months.length), rest = total - base * months.length;
    months.forEach(function (m, i) { plan[m] = base + (i < rest ? 1 : 0); });
    return plan;
  }
  // months that have fully ended on the demo date (15 June -> 5: January to May)
  function endedMonths(y, demo) {
    var d = dnum(demo), n = 0;
    for (var m = 0; m < 12; m++) if (d >= Date.UTC(y, m + 1, 1) / 864e5) n = m + 1;
    return n;
  }
  // the last month that has started on the demo date (-1 before the year)
  function lastStarted(y, demo) {
    var d = dnum(demo), last = -1;
    for (var m = 0; m < 12; m++) if (d >= Date.UTC(y, m, 1) / 864e5) last = m;
    return last;
  }
  // `done` from a single "actual to date" (import, old data): a count walks the elapsed months of the period and gives each
  // min(plan, what is left) (anything left over goes to the last of them); a percentage sits in the last elapsed month
  function doneFrom(o, actual, y, demo) {
    var done = rep(null);
    if (!finite(actual)) return done;
    var inP = monthsIn(o.start, o.end, y), last = lastStarted(y, demo), months = [];
    for (var m = 0; m <= last; m++) if (inP[m]) months.push(m);
    if (!months.length) { var f = inP.indexOf(true); months = [f >= 0 ? f : Math.max(0, Math.min(11, last))]; }
    var tail = months[months.length - 1];
    if (isPct(o.unit)) { done[tail] = actual; return done; }
    var left = actual, plan = o.plan || rep(0);
    months.forEach(function (mm, i) {
      if (left <= 1e-9 && i > 0) return;
      var give = Math.max(0, Math.min(finite(plan[mm]) ? plan[mm] : 0, left));
      done[mm] = give; left -= give;
    });
    if (left > 1e-9) done[tail] = (done[tail] || 0) + left;
    return done;
  }
  // fills the monthly fields of a KPI (a line's own KPI or an extra KPI) and drops the stored `actual`
  function normMonthly(o, y, demo) {
    if (Array.isArray(o.plan) && o.plan.length === 12) o.plan = o.plan.map(function (v) { return finite(v) ? v : 0; });
    else o.plan = planFor(o, y);
    if (Array.isArray(o.done) && o.done.length === 12) o.done = o.done.map(function (v) { return finite(v) ? v : null; });
    else o.done = doneFrom(o, o.actual === '' || o.actual === null || o.actual === undefined ? null : +o.actual, y, demo);
    delete o.actual;
    if (typeof o.kpiFull !== 'string') o.kpiFull = '';   // SPEC9 2.1: the KPI in full words ('' = composed from the numbers)
    if (PROOFS.indexOf(o.proof) < 0) o.proof = 'None';
    ANSWERS.forEach(function (a) { o[a] = o[a] === true; });
    return o;
  }
  // fills every line field with its default (SPEC5 1); used on load, on import and for new lines
  function normLine(l, y, demo) {
    l.kpis = Array.isArray(l.kpis) ? l.kpis : [];
    if (PAID_HOW.indexOf(l.paidHow) < 0) l.paidHow = '';
    if (BUCKETS.indexOf(l.bucket) < 0) l.bucket = '';
    if (!finite(l.paid)) l.paid = null;
    ['clause', 'justification', 'why'].forEach(function (f) { if (typeof l[f] !== 'string') l[f] = ''; });
    if (ACTIONS.indexOf(l.action) < 0) l.action = 'Keep';
    if (typeof l.priorityDecided !== 'boolean') l.priorityDecided = !!l.priorityId;
    // SPEC5b A: fixed (the amount is typed) or variable (the rate is typed, the amount follows; see TTS.recalc)
    if (BASES.indexOf(l.basis) < 0) l.basis = 'Fixed';
    l.rate = finite(l.rate) ? r2(l.rate) : null;
    normMonthly(l, y, demo);
    l.kpis.forEach(function (k) { normMonthly(k, y, demo); });
    return l;
  }

  // L(id, name, amount, paidHow, bucket, type, priority, counterpart, target, unit, start, end, source, owner, paid, propAmount, propType, propPriority, kpis)
  function L(id, name, amount, how, bucket, type, pri, cp, target, unit, start, end, source, owner, paid, pa, pt, pp, kpis) {
    return {
      id: id, name: name, amount: amount, paidHow: how, bucket: bucket, type: type, priorityId: pri,
      propAmount: pa, propType: pt, propPriorityId: pp,
      counterpart: cp, target: target, unit: unit, start: start, end: end, source: source, owner: owner, paid: paid,
      clause: '', justification: '', action: 'Keep', why: '', priorityDecided: true, basis: 'Fixed', rate: null,
      plan: null, done: null, proof: 'None', tracked: false, checked: false, due: false, paidFull: false,
      kpis: kpis || []
    };
  }
  // an unconditional line with no counterpart and no bucket, today and proposed
  function U(id, name, amount, how, source, owner, pa) {
    return L(id, name, amount, how, '', 'unconditional', null, '', null, '', null, null, source, owner, null, pa, 'unconditional', null);
  }
  function K(id, name, target, unit, start, end, source, owner) {
    return { id: id, name: name, target: target, unit: unit, start: start, end: end, source: source, owner: owner,
      plan: null, done: null, proof: 'None', tracked: false, checked: false, due: false, paidFull: false };
  }
  var J1 = '2027-01-01', D31 = '2027-12-31';

  // contract wording, justification, proposed action and why, per line (SPEC5 2.2 and 2.4)
  var EX_TEXT = {
    a1: ['The supplier grants a discount of 5.30% on the invoice on all products delivered during the agreement year.', 'None stated. Carried over from earlier agreements.',
      'Reduce', 'Move SEK 1.8 m to lines the customer earns by delivering.'],
    a2: ['A discount of 0.96% of net invoice sales is granted on invoices paid within 30 days.', 'Early payment.', 'Keep', 'It pays for paying on time.'],
    a3: ['A bonus of 1.72% of net invoice sales is paid after the end of the year.', 'None stated.',
      'Make conditional', 'Pay it for checkout coolers with single serve, not for the year ending.'],
    a4: ['The supplier pays SEK 250,000 for each leaflet week in which its products are featured with a display, up to 10 weeks a year.', 'Leaflet space and a display in store.',
      'Keep', 'It already has a number, a date and an owner.'],
    a5: ['The supplier contributes SEK 0.8 million to joint marketing activities in the regions, to be agreed during the year.', 'Regional store activations. No number agreed.',
      'Strengthen the counterpart', 'Agree the number of stores and link it to single serve.'],
    a6: ['The supplier pays SEK 1.4 million for the listing of up to six new packs in the central range by the end of April.', 'Central listing and distribution to the stores.',
      'Invest more', 'New packs are a priority and the counterpart is clear.'],
    a7: ['The supplier pays SEK 2.5 million for the zero sugar planogram to be in place in at least 90% of stores from February.', 'Shelf space for zero sugar, checked in the shelf audit.',
      'Invest more', 'More shelf for zero sugar, measured in the shelf audit.'],
    a8: ['An allowance of 0.72% of net invoice sales is paid when at least 85% of orders are placed in full pallets.', 'Lower handling cost for both parties.',
      'Keep', 'A fair exchange for full pallets, even if it serves no priority.'],
    b1: ['The supplier grants a discount of 5.18% on the invoice on all products delivered to the members during the agreement year.', 'None stated. Carried over from earlier agreements.',
      'Reduce', 'Move SEK 1.6 m to lines the members earn by delivering.'],
    b2: ['A central service fee of 1.43% of net invoice sales is paid to the group every quarter.', 'Central purchasing and invoicing. Nothing to deliver in return.',
      'Keep', 'It pays for the central services the group provides.'],
    b3: ['A discount of 0.92% of net invoice sales is granted on invoices paid within 30 days.', 'Early payment.', 'Keep', 'It pays for paying on time.'],
    b4: ['A bonus of 1.24% of net invoice sales is paid after the end of the year.', 'None stated.',
      'Reduce', 'Paid whatever happens: move SEK 0.9 m to new stores and listings.'],
    b5: ['The supplier pays SEK 300,000 for each central leaflet week in which its products are featured with a secondary display, up to 12 weeks a year.', 'Leaflet space and a secondary display in the member stores.',
      'Keep', 'It already has a number, a date and an owner.'],
    b6: ['The supplier contributes SEK 2.1 million to the regional member fund for a stand at up to six regional trade fairs by the end of November.', 'Our stand at the regional trade fairs.',
      'Move to a priority', 'Use the trade fairs to show zero sugar.'],
    b7: ['The supplier pays SEK 75,000 for each new store that opens with a checkout cooler, up to eight stores a year.', 'A checkout cooler in every new store.',
      'Invest more', 'Every new store should open with single serve at the checkout.'],
    b8: ['The supplier pays SEK 1.6 million for the listing of up to five new packs in the member range by the end of May.', 'Listing in the member range.',
      'Invest more', 'New packs are a priority and the counterpart is clear.'],
    b9: ['The supplier pays SEK 125,000 for each week with a zero sugar end cap, up to 20 weeks between March and October.', 'End-cap space for zero sugar, checked in the shelf audit.',
      'Invest more', 'End caps are the best place to grow zero sugar.'],
    b10: ['An allowance of 0.24% of net invoice sales covers damaged and returned products.', 'None stated. Covers damage in transport and in store.',
      'Fold into the list price', 'A cost of doing business, not a trade term.'],
    c1: ['The supplier grants a discount of 3.41% on the invoice on all products delivered during the agreement year.', 'None stated. Carried over from earlier agreements.',
      'Reduce', 'Move SEK 0.5 m to coolers and zero sugar.'],
    c2: ['A discount of 1.73% of net invoice sales is granted on invoices paid within 20 days.', 'Early payment.', 'Keep', 'It pays for paying on time.'],
    c3: ['The supplier pays SEK 1.1 million for the listing of six new packs in the wholesale range by the end of March.', 'Listing in the wholesale range and its catalogue.',
      'Keep', 'Delivered in full.'],
    c4: ['The supplier pays SEK 425,000 for each catalogue feature of its products, up to eight features a year.', 'Catalogue space for our products.',
      'Reduce', 'Fewer features, each planned three months ahead.'],
    c5: ['The supplier pays SEK 18,000 for each cooler placed in an outlet the wholesaler supplies, up to 150 coolers a year.', 'Coolers in the outlets, checked on outlet visits.',
      'Invest more', 'Coolers are how single serve grows in the outlets.'],
    c6: ['The supplier contributes SEK 2.0 million to bring the zero sugar range to the outlets the wholesaler supplies during the year.', 'Zero sugar on offer in the outlets, measured in the sell-out data.',
      'Invest more', 'Zero sugar is still missing in three outlets out of ten.'],
    c7: ['The supplier pays SEK 400,000 for a monthly file of sell-out data by outlet, twelve files a year.', 'Sell-out data by outlet.',
      'Keep', 'The files are how we measure the other lines.'],
    c8: ['An allowance of 1.44% of net invoice sales is paid when at least 90% of orders are placed by EDI in full pallets.', 'Lower handling cost for both parties.',
      'Keep', 'A fair exchange for EDI orders in full pallets.']
  };
  // month by month (SPEC5 2.5): [plan, done, proof, tracked / checked / due / paid in full]; _ = nothing recorded
  var _ = null;
  var EX_MONTHS = {
    a4: [[0, 1, 1, 1, 1, 1, 0, 1, 1, 1, 1, 1], rep(_), 'Leaflet copy', 'nnyn'],
    a6: [[2, 2, 1, 1, 0, 0, 0, 0, 0, 0, 0, 0], [2, 1, 1, 1, _, _, _, _, _, _, _, _], 'Listing file', 'yyyy'],
    a6k1: [[0, 40, 50, 60, 70, 80, 0, 0, 0, 0, 0, 0], [_, 40, 52, 58, 64, _, _, _, _, _, _, _], 'Customer report', 'ynyy'],
    a7: [[0, 90, 90, 90, 90, 90, 90, 90, 90, 90, 90, 90], [_, 91, 92, 93, 93, _, _, _, _, _, _, _], 'Shelf audit', 'yynn'],
    a8: [rep(85), [86, 85, 87, 88, 87, _, _, _, _, _, _, _], 'Customer report', 'yynn'],
    b5: [rep(1), [1, 1, 1, 1, 1, 1, _, _, _, _, _, _], 'Leaflet copy', 'yynn'],
    b6: [[0, 1, 0, 1, 1, 0, 0, 1, 1, 1, 0, 0], [_, 1, _, 1, 0, _, _, _, _, _, _, _], 'Customer report', 'ynyn'],
    b7: [[0, 1, 1, 1, 1, 0, 1, 1, 1, 1, 0, 0], [_, 1, 1, 1, 1, _, _, _, _, _, _, _], 'Store visit photo', 'yynn'],
    b7k1: [[0, 0, 70, 80, 90, 90, 90, 90, 90, 90, 90, 90], [_, _, 70, 72, 75, _, _, _, _, _, _, _], 'Store visit photo', 'ynnn'],
    b8: [[1, 1, 1, 1, 1, 0, 0, 0, 0, 0, 0, 0], [1, 1, 0, 1, 0, _, _, _, _, _, _, _], 'Listing file', 'yyyy'],
    b9: [[0, 0, 2, 3, 3, 2, 2, 3, 3, 2, 0, 0], rep(_), 'None', 'nnnn'],
    c3: [[2, 2, 2, 0, 0, 0, 0, 0, 0, 0, 0, 0], [2, 2, 2, _, _, _, _, _, _, _, _, _], 'Listing file', 'yyyy'],
    c4: [[1, 0, 1, 1, 1, 0, 1, 1, 0, 1, 1, 0], [1, _, 1, 1, 0, _, _, _, _, _, _, _], 'Leaflet copy', 'ynyn'],
    c5: [[10, 10, 15, 15, 15, 15, 15, 15, 15, 10, 10, 5], [10, 10, 12, 11, 15, _, _, _, _, _, _, _], 'Store visit photo', 'yynn'],
    c5k1: [[0, 0, 95, 95, 95, 95, 95, 95, 95, 95, 95, 95], [_, _, 95, 96, 96, _, _, _, _, _, _, _], 'Store visit photo', 'ynnn'],
    c6: [rep(70), [70, 70, 71, 72, 73, _, _, _, _, _, _, _], 'Sell-out data', 'yynn'],
    c7: [rep(1), rep(_), 'None', 'nnnn'],
    c8: [rep(90), [90, 91, 91, 92, 91, _, _, _, _, _, _, _], 'Customer report', 'yynn']
  };
  // SPEC9 8.1: each KPI in full, specific words (kpiFull): the pack or the action, the measure, the threshold and the scope
  var EX_KPI = {
    a4: '10 leaflet weeks, each with a floor display of the 1.5 litre bottle in all large stores',
    a6: '6 new packs, including the multipack of 8 slim cans, listed in the central range',
    a6k1: 'Multipack of 8 slim cans listed in stores covering at least 80% of weighted distribution',
    a7: 'Zero sugar planogram implemented in at least 90% of stores',
    a8: 'At least 85% of orders to the central warehouse placed in full pallets',
    b5: '12 central leaflet weeks, each with a secondary display of the 6-pack of 33 cl cans in all member stores',
    b6: '6 regional trade fairs with our stand and a tasting of the new packs for member store owners',
    b7: '8 new stores opened with a checkout cooler stocked with 33 cl cans and 50 cl bottles',
    b7k1: 'Full single-serve range of five packs on shelf in at least 90% of the new stores',
    b8: '5 new packs, including the 4-pack of 25 cl cans, listed in the member range',
    b9: '20 end-cap weeks for the zero sugar 1.5 litre bottle in all large member stores',
    c3: '6 new packs, including the 33 cl glass bottle, listed in the wholesale range and its catalogue',
    c4: '8 catalogue features of the 33 cl glass bottle, each on a full page and agreed three months ahead',
    c5: '150 coolers placed in outlets the wholesaler supplies, each confirmed by a photo from the outlet visit',
    c5k1: 'At least 95% of the placed coolers stocked with single-serve 33 cl cans and 50 cl bottles',
    c6: 'Zero sugar 33 cl can and 50 cl bottle on offer in at least 70% of the outlets the wholesaler supplies',
    c7: '12 monthly sell-out files by outlet and by pack, each received by the 10th of the following month',
    c8: 'At least 90% of orders placed by EDI and in full pallets'
  };

  // SPEC5b E: the variable lines and their rates (they reproduce today's amounts exactly); every other line is Fixed
  var EX_RATES = { a1: 5.30, a2: 0.96, a3: 1.72, a8: 0.72, b1: 5.18, b2: 1.43, b3: 0.92, b4: 1.24, b10: 0.24, c1: 3.41, c2: 1.73, c8: 1.44 };
  // SPEC5b E: promotional spend (SEK m) and pack groups [name, share of gross revenue %, promotional spend SEK m]
  var EX_PROMO = { a: 13.2, b: 18.5, c: 4.3 };
  var EX_PACKS = {
    a: [['Single serve', 22, 1.2], ['Multipacks', 43, 7.4], ['Large bottles', 35, 4.6]],
    b: [['Single serve', 20, 1.5], ['Multipacks', 45, 10.5], ['Large bottles', 35, 6.5]],
    c: [['Single serve', 55, 1.6], ['Multipacks', 15, 0.9], ['Large bottles', 30, 1.8]]
  };

  function exampleState() {
    var C = 'conditional', UN = 'unconditional', ON = 'On invoice', OFF = 'Off invoice', FIX = 'Fixed amount';
    var LI = 'Listings', SH = 'Shelf space', IS = 'In-store marketing', PR = 'Promotions', EF = 'Efficiency';
    var YEAR = 2027, DEMO = '2027-06-15';
    function cust(id, name, channel, listSales, rLY, iLY, lines, mustWin) {
      lines.forEach(function (l) {
        var t = EX_TEXT[l.id];
        if (t) { l.clause = t[0]; l.justification = t[1]; l.action = t[2]; l.why = t[3]; }
        if (EX_RATES[l.id] !== undefined) { l.basis = 'Variable'; l.rate = EX_RATES[l.id]; }
        [l].concat(l.kpis).forEach(function (o) {
          var mm = EX_MONTHS[o.id];
          if (mm) {
            o.plan = mm[0].slice(); o.done = mm[1].slice(); o.proof = mm[2];
            ANSWERS.forEach(function (a, i) { o[a] = mm[3].charAt(i) === 'y'; });
          }
          if (EX_KPI[o.id]) o.kpiFull = EX_KPI[o.id];
        });
        normLine(l, YEAR, DEMO);
      });
      return { id: id, name: name, channel: channel, example: true, sourceNote: 'Example data', listSales: listSales, revenueLY: rLY,
        investmentLY: iLY, promoSpend: EX_PROMO[id] || 0,
        packs: (EX_PACKS[id] || []).map(function (p) { return { name: p[0], share: p[1], promo: p[2] }; }),
        mustWinId: mustWin || null, scriptEdits: {}, lines: lines };
    }
    var a5 = L('a5', 'Marketing support', 0.8, FIX, IS, C, null, 'Stores activated with single serve in the regions', null, 'stores', J1, D31, 'Store visit report', 'Trade marketing', null, 1.5, C, 'p1');
    a5.priorityDecided = false;   // the priority is still open: the tool suggests Drive single serve (SPEC5 2.3)
    return {
      version: 5, year: YEAR, demoDate: DEMO,
      priorities: [
        { id: 'p1', name: 'Drive single serve', winning: 'Chilled single serve at every checkout and every food-to-go spot.' },
        { id: 'p2', name: 'Grow zero sugar', winning: 'Zero sugar on shelf, cold and in every promotion.' },
        { id: 'p3', name: 'Land new packs', winning: 'New packs listed and on shelf in most stores within eight weeks of launch.' },
        { id: 'p4', name: 'Fewer, better promotions', winning: 'Fewer promo weeks, each with a display and a plan agreed in advance.' }
      ],
      customers: [
        cust('a', 'Customer A', 'Home', 220.7, 203.0, 25.5, [
          U('a1', 'Agreement discount', 11.7, ON, 'Price agreement', 'KAM', 9.9),
          U('a2', 'Payment terms discount', 2.0, OFF, 'Finance system', 'Finance controller', 2.0),
          L('a3', 'Year-end bonus', 3.6, OFF, IS, UN, null, 'Checkout coolers with single serve', 60, 'stores', '2027-03-01', D31, 'Store visit report', 'Field sales lead', null, 3.6, C, 'p1'),
          L('a4', 'Leaflet allowance', 2.5, FIX, PR, C, 'p4', 'Leaflet weeks with a display', 10, 'weeks', J1, D31, 'Promotion calendar', 'KAM', 1.5, 2.5, C, 'p4'),
          a5,
          L('a6', 'Listing fee for new packs', 1.4, FIX, LI, C, 'p3', 'New packs listed in the central range', 6, 'SKUs', J1, '2027-04-30', 'Customer listing file', 'Category lead', 1.4, 1.9, C, 'p3', [
            K('a6k1', 'Distribution of the multipack of 8 slim cans', 80, '% of weighted distribution', '2027-02-01', '2027-06-30', 'Store distribution report', 'Category lead')]),
          L('a7', 'Shelf space', 2.5, FIX, SH, C, 'p2', 'Zero sugar planogram implemented', 90, '% of stores', '2027-02-01', D31, 'Shelf audit', 'Field sales lead', 1.0, 3.1, C, 'p2'),
          L('a8', 'Logistics allowance', 1.5, OFF, EF, C, null, 'Orders in full pallets', 85, '%', J1, D31, 'Order system', 'Supply chain planner', 0.7, 1.5, C, null)
        ]),
        cust('b', 'Customer B', 'Home', 264.7, 246.0, 32.4, [
          U('b1', 'Agreement discount', 13.7, ON, 'Price agreement', 'KAM', 12.1),
          U('b2', 'Central service fee', 3.6, OFF, 'Finance system', 'Finance controller', 3.6),
          U('b3', 'Payment terms discount', 2.3, OFF, 'Finance system', 'Finance controller', 2.3),
          U('b4', 'Year-end bonus', 3.1, OFF, 'Finance system', 'KAM', 2.2),
          L('b5', 'Central leaflet allowance', 3.6, FIX, PR, C, 'p4', 'Leaflet weeks with a secondary display', 12, 'weeks', J1, D31, 'Promotion calendar', 'KAM', 1.8, 3.6, C, 'p4'),
          L('b6', 'Regional member fund', 2.1, FIX, IS, C, null, 'Regional trade fairs with our stand', 6, 'events', J1, '2027-11-30', 'Event list', 'Trade marketing', 1.4, 2.1, C, 'p2'),
          L('b7', 'New-store opening fund', 0.6, FIX, IS, C, 'p1', 'New stores opened with a checkout cooler', 8, 'stores', J1, D31, 'Store visit report', 'Field sales lead', 0.2, 1.9, C, 'p1', [
            K('b7k1', 'New stores with the full single-serve range', 90, '% of new stores', J1, D31, 'Store visit report', 'Field sales lead')]),
          L('b8', 'Listing fee', 1.6, FIX, LI, C, 'p3', 'New packs listed in the member range', 5, 'SKUs', J1, '2027-05-31', 'Customer listing file', 'Category lead', 1.6, 2.2, C, 'p3'),
          L('b9', 'Zero sugar end caps', 2.5, FIX, SH, C, 'p2', 'End-cap weeks for zero sugar', 20, 'weeks', '2027-03-01', '2027-10-31', 'Shelf audit', 'Field sales lead', null, 3.1, C, 'p2'),
          U('b10', 'Damage and returns allowance', 0.6, OFF, 'Claims log', 'Customer service lead', 0.0)
        ]),
        cust('c', 'Customer C', 'Away from Home', 143.9, 128.7, 18.0, [
          U('c1', 'Agreement discount', 4.9, ON, 'Price agreement', 'KAM', 4.4),
          U('c2', 'Payment terms discount', 2.4, OFF, 'Finance system', 'Finance controller', 2.4),
          L('c3', 'Range listing support', 1.1, FIX, LI, C, 'p3', 'New packs in the wholesale range', 6, 'SKUs', J1, '2027-03-31', 'Customer listing file', 'KAM', 1.1, 1.1, C, 'p3'),
          L('c4', 'Catalogue features', 3.4, FIX, PR, C, 'p4', 'Catalogue features planned three months ahead', 8, 'features', J1, D31, 'Promotion calendar', 'KAM', 1.7, 2.7, C, 'p4'),
          L('c5', 'Cooler placement', 2.7, FIX, IS, C, 'p1', 'Coolers placed in outlets the wholesaler supplies', 150, 'outlets', J1, D31, 'Outlet visit report', 'Field sales lead', 0.9, 3.4, C, 'p1', [
            K('c5k1', 'Coolers stocked with single serve', 95, '% of coolers', J1, D31, 'Outlet visit report', 'Field sales lead')]),
          L('c6', 'Zero sugar range support', 2.0, FIX, LI, C, 'p2', 'Outlets with zero sugar on offer', 70, '% of outlets', J1, D31, 'Sell-out data', 'Revenue manager', 1.0, 2.5, C, 'p2'),
          L('c7', 'Sell-out data sharing', 0.4, FIX, EF, C, null, 'Monthly sell-out files by outlet', 12, 'files', J1, D31, 'Customer data portal', 'Revenue manager', null, 0.4, C, null),
          L('c8', 'Logistics allowance', 2.0, OFF, EF, C, null, 'Orders by EDI in full pallets', 90, '%', J1, D31, 'Order system', 'Supply chain planner', 0.9, 2.0, C, null)
        ])
      ],
      ui: { screen: 'priorities', scope: 'a', agreementMode: 'today', lastCustomerId: 'a' }
    };
  }
  TTS.exampleState = exampleState;

  // ---------------------------------------------------------------- formatting
  var MONTH = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
  function minus(s) { return String(s).replace('-', '−'); }
  function r1(n) { return Math.round((+n || 0) * 10 + 1e-9) / 10; }

  var fmt = {
    m: function (n) { var s = r1(n).toFixed(1); if (s === '-0.0') s = '0.0'; return minus(s); },
    money: function (n) { return 'SEK ' + fmt.m(n) + ' m'; },
    pct: function (share) { return Math.round((+share || 0) * 100) + '%'; },
    pct1: function (x) { var s = ((+x || 0) * 100).toFixed(1); if (s === '-0.0') s = '0.0'; return minus(s) + '%'; },
    pct1s: function (x) { var s = fmt.pct1(x); return (x > 0 && s !== '0.0%') ? '+' + s : s; },
    delta: function (n) { var v = r1(n); if (v === 0) return '0.0'; return (v > 0 ? '+' : '') + fmt.m(v); },
    // a rate already in % (1.72 = 1.72%): two decimals, trailing zeros trimmed to at least one ("5.3%", "1.72%"); '' when none
    rate: function (x) { if (!finite(x)) return ''; var s = r2(x).toFixed(2).replace(/0$/, ''); if (s === '-0.0') s = '0.0'; return minus(s) + '%'; },
    num: function (n) {
      if (!finite(n)) return '';
      return minus(String(Math.round(n * 100) / 100));
    },
    parse: function (v) {
      if (v === null || v === undefined) return null;
      if (typeof v === 'number') return isFinite(v) ? v : NaN;
      var s = String(v).replace(/[\s ]/g, '').replace('−', '-').replace(',', '.');
      if (s === '') return null;
      if (!/^[-+]?(\d+\.?\d*|\.\d+)$/.test(s)) return NaN;
      return parseFloat(s);
    },
    target: function (target, unit) {
      if (!finite(target)) return '';
      var u = String(unit || '').trim();
      if (u.charAt(0) === '%') return fmt.num(target) + u;
      return u ? fmt.num(target) + ' ' + u : fmt.num(target);
    },
    date: function (iso) {
      if (!iso) return '';
      var p = String(iso).split('-');
      return (+p[2]) + ' ' + MON[+p[1] - 1] + ' ' + p[0];
    },
    isoWeek: function (iso) {
      var p = String(iso).split('-');
      var d = new Date(Date.UTC(+p[0], +p[1] - 1, +p[2]));
      var day = d.getUTCDay() || 7;
      d.setUTCDate(d.getUTCDate() + 4 - day);
      var y0 = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
      return Math.ceil(((d - y0) / 864e5 + 1) / 7);
    },
    period: function (start, end) {
      if (!end) return '';
      var y = (TTS.state && TTS.state.year) || 2027;
      var s = start || (y + '-01-01');
      var jan1 = s === y + '-01-01', dec31 = end === y + '-12-31';
      if (jan1 && dec31) return 'Whole year';
      if (jan1) return 'By week ' + fmt.isoWeek(end);
      if (dec31) return 'From ' + MONTH[+s.split('-')[1] - 1];
      return MON[+s.split('-')[1] - 1] + ' to ' + MON[+end.split('-')[1] - 1];
    }
  };
  TTS.fmt = fmt;

  // ---------------------------------------------------------------- engine
  function S() { return TTS.state; }
  function dayNum(iso) { var p = String(iso).split('-'); return Date.UTC(+p[0], +p[1] - 1, +p[2]) / 864e5; }

  var E = {};
  E.ACTIONS = ACTIONS;
  E.PROOFS = PROOFS;
  E.MONTHS = MON;
  E.priority = function (id) {
    if (!id) return null;
    var ps = S().priorities;
    for (var i = 0; i < ps.length; i++) if (ps[i].id === id) return ps[i];
    return null;
  };
  E.amount = function (line, mode) { var v = mode === 'proposed' ? line.propAmount : line.amount; return finite(v) ? v : 0; };
  E.type = function (line, mode) { return (mode === 'proposed' ? line.propType : line.type) === 'conditional' ? 'conditional' : 'unconditional'; };
  E.priorityId = function (line, mode) { var id = mode === 'proposed' ? line.propPriorityId : line.priorityId; return E.priority(id) ? id : null; };
  // counterpart, target number and date (the owner is checked separately, see classOf)
  E.hasNumber = function (line) { return !blank(line.counterpart) && finite(line.target) && !!line.end; };
  E.kpiComplete = function (k) { return !blank(k.name) && finite(k.target) && !!k.end; };
  // what a line (or, with nameField 'name', an extra KPI) still lacks: counterpart, target, date and, for a line, owner
  E.missing = function (o, nameField) {
    var out = [];
    if (blank(o[nameField || 'counterpart'])) out.push('counterpart');
    if (!finite(o.target)) out.push('target');
    if (!o.end) out.push('date');
    if (!nameField && blank(o.owner)) out.push('owner');
    return out;
  };
  // SPEC5 3.3: unconditional when the type is; conditional with a counterpart, a target number, a date and an owner;
  // otherwise partially conditional (never counted as conditional; sits in the Unconditional bucket)
  E.classOf = function (line, mode) {
    if (E.type(line, mode) === 'unconditional') return 'unconditional';
    return E.hasNumber(line) && !blank(line.owner) ? 'conditional' : 'partial';
  };
  E.CLASS_LABEL = { unconditional: 'Unconditional', partial: 'Partially conditional', conditional: 'Conditional' };
  E.lineStatus = E.classOf;   // older name, same values
  // counterpart strength (SMART) of a conditional-type line; an unconditional line has no counterpart: score 0
  E.smart = function (line, mode) {
    var cond = E.type(line, mode) === 'conditional';
    var s = { specific: cond && !blank(line.counterpart), measurable: cond && finite(line.target), owner: cond && !blank(line.owner), timebound: cond && !!line.end };
    s.score = (s.specific ? 1 : 0) + (s.measurable ? 1 : 0) + (s.owner ? 1 : 0) + (s.timebound ? 1 : 0);
    return s;
  };
  E.aligned = function (line, mode) { return !!E.priorityId(line, mode); };
  // suggested priority (SPEC5 3.3): only for a conditional-type line with no priority that nobody has decided yet
  var SUG_IGNORE = ['drive', 'grow', 'land', 'fewer', 'better', 'more', 'with', 'and', 'the', 'new'];
  var SUG_HINTS = { 'Listings': 'listing pack range', 'Shelf space': 'shelf', 'Promotions': 'promotion leaflet', 'In-store marketing': 'store display cooler', 'Efficiency': '' };
  function sugWords(s) {
    return String(s || '').toLowerCase().split(/[^a-z0-9]+/).filter(Boolean)
      .map(function (w) { return w.length > 3 && /[^s]s$/.test(w) ? w.slice(0, -1) : w; });
  }
  E.suggestPriority = function (line) {
    if (!line || E.type(line, 'today') !== 'conditional' || E.priorityId(line, 'today') || line.priorityDecided !== false) return null;
    var text = sugWords(line.name + ' ' + (line.counterpart || '') + ' ' + (SUG_HINTS[line.bucket] || ''));
    var best = null, bestScore = 0;
    S().priorities.forEach(function (p) {
      var seen = {}, score = 0;
      sugWords(p.name).forEach(function (w) {
        if (SUG_IGNORE.indexOf(w) >= 0 || seen[w]) return;
        seen[w] = 1;
        if (text.indexOf(w) >= 0) score += 1;
      });
      if (score > bestScore) { best = p.id; bestScore = score; }
    });
    return best;
  };
  E.summary = function (lines, mode) {
    var s = { total: 0, unconditional: 0, partial: 0, conditional: 0, linked: 0 };
    (lines || []).forEach(function (l) {
      var a = E.amount(l, mode);
      s.total += a;
      s[E.classOf(l, mode)] += a;
      if (E.priorityId(l, mode)) s.linked += a;
    });
    var t = s.total > 0 ? s.total : 0;
    s.shareUnconditional = t ? s.unconditional / t : 0;
    s.sharePartial = t ? s.partial / t : 0;
    s.shareConditional = t ? s.conditional / t : 0;
    s.shareLinked = t ? s.linked / t : 0;
    s.nonumber = s.partial; s.shareNoNumber = s.sharePartial;   // older names, same values
    return s;
  };
  // amount and share of the spend linked to a priority (whatever the line's type)
  E.linkedShare = function (lines, mode) {
    var s = E.summary(lines, mode);
    return { amount: s.linked, total: s.total, share: s.shareLinked };
  };
  // the price waterfall in numbers, of a customer (or the sum of an array of customers) in a view (SPEC5b B):
  // gross revenue at list price - on invoice = first net - off invoice = second net - promotional spend = third net
  E.sales = function (cs, mode) {
    var list = Array.isArray(cs) ? cs : [cs];
    var o = { gross: 0, onInvoice: 0, firstNet: 0, offInvoice: 0, secondNet: 0, promo: 0, thirdNet: 0, tradeTerms: 0 };
    list.forEach(function (c) {
      if (!c) return;
      var g = finite(c.listSales) ? c.listSales : 0;
      (c.lines || []).forEach(function (l) {
        var a = E.amount(l, mode);
        if (mode === 'proposed' && l.action === 'Fold into the list price') g -= E.amount(l, 'today');
        if (l.paidHow === 'On invoice') o.onInvoice += a; else o.offInvoice += a;
      });
      o.gross += g;
      o.promo += finite(c.promoSpend) ? c.promoSpend : 0;
    });
    o.firstNet = o.gross - o.onInvoice;
    o.secondNet = o.firstNet - o.offInvoice;
    o.thirdNet = o.secondNet - o.promo;
    o.tradeTerms = o.onInvoice + o.offInvoice;
    return o;
  };
  // "% of gross revenue" of an amount: a fraction of the gross revenue of the same view (0 when gross revenue <= 0)
  E.pctOfGross = function (amount, cs, mode) { var g = E.sales(cs, mode).gross; return g > 0 ? (+amount || 0) / g : 0; };
  // the contract base of a line's rate: gross revenue for an on-invoice line, the first net for every other line
  E.rateBase = function (line) { return line && line.paidHow === 'On invoice' ? 'gross' : 'firstNet'; };
  // the rate of a line in % of its base (1.72 = 1.72%), two decimals: the stored rate of a variable line; for a fixed line
  // (or a variable line with no rate yet) today's amount ÷ today's base. null when the base is not above 0.
  E.rateOf = function (line, c) {
    if (!line) return null;
    if (line.basis === 'Variable' && finite(line.rate)) return line.rate;
    var base = c ? E.sales(c, 'today')[E.rateBase(line)] : 0;
    return base > 0 ? r2(E.amount(line, 'today') / base * 100) : null;
  };
  E.spendByPriority = function (lines, mode) {
    var total = 0, by = {}, none = 0;
    (lines || []).forEach(function (l) {
      var a = E.amount(l, mode), p = E.priorityId(l, mode);
      total += a;
      if (p) by[p] = (by[p] || 0) + a; else none += a;
    });
    var out = S().priorities.map(function (p) {
      var a = by[p.id] || 0;
      return { priorityId: p.id, name: p.name, amount: a, share: total > 0 ? a / total : 0 };
    });
    out.push({ priorityId: null, name: 'No priority', amount: none, share: total > 0 ? none / total : 0 });
    return out;
  };
  // buckets in their fixed order (SPEC4 1); "No bucket" is last and only shown when above 0
  E.BUCKETS = [
    { key: 'unconditional', name: 'Unconditional', color: '#AAB3C4' },
    { key: 'listings', name: 'Listings', color: '#FBCB1F' },
    { key: 'shelf', name: 'Shelf space', color: '#28A878' },
    { key: 'instore', name: 'In-store marketing', color: '#8FDDBD' },
    { key: 'promotions', name: 'Promotions', color: '#48C9FF' },
    { key: 'efficiency', name: 'Efficiency', color: '#232341' },
    { key: 'nobucket', name: 'No bucket', color: '#E6E9EF' }
  ];
  // where a line's money counts in a view: Unconditional (unconditional or partially conditional), else its bucket, else No bucket
  E.bucketOf = function (line, mode) {
    if (E.classOf(line, mode) !== 'conditional') return 'Unconditional';
    return BUCKETS.indexOf(line.bucket) >= 0 ? line.bucket : 'No bucket';
  };
  E.spendByBucket = function (lines, mode) {
    var by = {}, total = 0;
    (lines || []).forEach(function (l) {
      var a = E.amount(l, mode), b = E.bucketOf(l, mode);
      total += a;
      by[b] = (by[b] || 0) + a;
    });
    return E.BUCKETS.map(function (b) {
      var a = by[b.name] || 0;
      return { key: b.key, name: b.name, color: b.color, amount: a, share: total > 0 ? a / total : 0 };
    });
  };
  // THE price structure of a customer in a view (SPEC5c 1): the Contract analysis table AND the Agreement today picture are both
  // drawn from these rows (same rows, same order, same numbers). Every row has the same keys:
  // { kind: 'anchor'|'line'|'promo', key ('gross'|'firstNet'|'secondNet'|'thirdNet' on anchors, 'promo' on the promo row, null on
  //   lines), block ('on'|'off' on lines, null elsewhere), label, hint (text or null), amount (SEK m, positive; a deduction is shown
  //   negative), pct (amount ÷ gross revenue of the view), lineId, bucket, bucketKey, cls, basis, rate (line rows; null elsewhere) }
  var WF_HINT = {
    firstNet: 'After the on-invoice contract points',
    secondNet: 'After the off-invoice contract points: rebates, bonuses and fixed allowances',
    promo: 'Price discounts in promotions. Not a contract point.',
    thirdNet: 'After promotional spend'
  };
  E.waterfall = function (c, mode) {
    var sl = E.sales(c, mode), gross = sl.gross, rows = [];
    function row(kind, key, label, amount, line) {
      var b = line ? E.bucketOf(line, mode) : null, bk = null;
      if (b) E.BUCKETS.forEach(function (x) { if (x.name === b) bk = x.key; });
      rows.push({ kind: kind, key: key, block: line ? (line.paidHow === 'On invoice' ? 'on' : 'off') : null, label: label,
        hint: (!line && WF_HINT[key]) || null, amount: amount, pct: gross > 0 ? amount / gross : 0,
        lineId: line ? line.id : null, bucket: b, bucketKey: bk, cls: line ? E.classOf(line, mode) : null,
        basis: line ? (line.basis === 'Variable' ? 'Variable' : 'Fixed') : null, rate: line ? E.rateOf(line, c) : null });
    }
    function block(onInvoice) {
      (c.lines || []).forEach(function (l) { if ((l.paidHow === 'On invoice') === onInvoice) row('line', null, l.name, E.amount(l, mode), l); });
    }
    row('anchor', 'gross', 'Gross revenue at list price', sl.gross);
    block(true);
    row('anchor', 'firstNet', 'First net (net invoice sales)', sl.firstNet);
    block(false);
    row('anchor', 'secondNet', 'Second net', sl.secondNet);
    row('promo', 'promo', 'Promotional spend (estimate)', sl.promo);
    row('anchor', 'thirdNet', 'Third net (net-net)', sl.thirdNet);
    return rows;
  };
  // the waterfall by pack group (SPEC5b B): null without packs; else [All packs, then one column per pack group], each
  // { name, gross, onInvoice, firstNet, offInvoice, secondNet, promo, thirdNet, pct: { same keys, fractions of that column's gross } }.
  // Gross revenue, on invoice and off invoice are split by the pack group's share of gross revenue; the promotional spend is its own.
  var PACK_KEYS = ['gross', 'onInvoice', 'firstNet', 'offInvoice', 'secondNet', 'promo', 'thirdNet'];
  E.packWaterfall = function (c, mode) {
    if (!c || !Array.isArray(c.packs) || !c.packs.length) return null;
    var sl = E.sales(c, mode);
    function col(name, o) {
      o.name = name;
      o.firstNet = o.gross - o.onInvoice; o.secondNet = o.firstNet - o.offInvoice; o.thirdNet = o.secondNet - o.promo;
      o.pct = {};
      PACK_KEYS.forEach(function (k) { o.pct[k] = o.gross > 0 ? o[k] / o.gross : 0; });
      return o;
    }
    var out = [col('All packs', { gross: sl.gross, onInvoice: sl.onInvoice, offInvoice: sl.offInvoice, promo: sl.promo })];
    c.packs.forEach(function (p) {
      var s = (finite(p.share) ? p.share : 0) / 100;
      out.push(col(p.name, { gross: sl.gross * s, onInvoice: sl.onInvoice * s, offInvoice: sl.offInvoice * s, promo: finite(p.promo) ? p.promo : 0 }));
    });
    return out;
  };

  // ---- KPIs from the monthly rows (SPEC5 3.4)
  E.isPercent = isPct;
  E.periodMonths = function (start, end) { return monthsIn(start, end, S().year); };   // 12 booleans: months the period touches
  E.defaultPlan = function (o) { return planFor(o, S().year); };                      // 12 numbers for a KPI-like { target, unit, start, end }
  E.monthsEnded = function () { return endedMonths(S().year, S().demoDate); };         // 5 on 15 June: January to May have ended
  E.currentMonth = function () { return lastStarted(S().year, S().demoDate); };        // 5 on 15 June (0 = January; -1 before the year)
  E.doneFromActual = function (o, actual) { return doneFrom(o, actual, S().year, S().demoDate); };
  // actual to date of a KPI-like { unit, done }: a count sums what is recorded, a percentage takes the latest recorded month
  E.actualOf = function (o) {
    var d = o.done || [], sum = 0, latest = null, has = false;
    for (var m = 0; m < 12; m++) if (finite(d[m])) { has = true; sum += d[m]; latest = d[m]; }
    if (!has) return null;
    return isPct(o.unit) ? latest : sum;
  };
  E.expectedOf = function (o) {
    if (isPct(o.unit)) return null;
    var n = E.monthsEnded(), p = o.plan || [], sum = 0;
    for (var m = 0; m < n && m < 12; m++) sum += finite(p[m]) ? p[m] : 0;
    return sum;
  };
  E.elapsed = function (start, end) {
    if (!end) return 0;
    var st = S();
    var s = dayNum(start || (st.year + '-01-01')), e = dayNum(end), d = dayNum(st.demoDate);
    var total = e - s + 1;
    if (total <= 0) return d >= e ? 1 : 0;
    return Math.max(0, Math.min(1, (d - s + 1) / total));
  };
  E.statusLabel = function (s) { return { executed: 'Fully executed', partly: 'Partially executed', notexecuted: 'Not executed' }[s] || ''; };
  // the KPI in words (SPEC9 2.1): kpiFull when typed, otherwise composed from the name, the target and the unit
  function kpiText(name, k) {
    if (typeof k.kpiFull === 'string' && k.kpiFull.trim()) return k.kpiFull.trim();
    name = String(name || '');
    if (isPct(k.unit)) return name + ': ' + fmt.target(k.target, k.unit);
    return fmt.num(k.target) + ' ' + (/^[A-Z][a-z]/.test(name) ? name.charAt(0).toLowerCase() + name.slice(1) : name);
  }
  function mkKpi(c, line, k, isExtra) {
    var kind = isPct(k.unit) ? 'percent' : 'count';
    var actual = E.actualOf(k), expected = E.expectedOf(k), measured = actual !== null;
    var recorded = (k.done || []).reduce(function (s, v) { return s + (finite(v) ? v : 0); }, 0);
    // SPEC9 2.2: nothing recorded (or it adds up to 0) -> notexecuted; reached -> executed; anything in between -> partly
    var status = !measured || Math.abs(recorded) < 1e-9 ? 'notexecuted'
      : (kind === 'percent' ? actual >= k.target - 1e-9 : actual >= expected - 1e-9) ? 'executed' : 'partly';
    var delivered = !measured ? 0 : k.target > 0 ? Math.max(0, Math.min(actual / k.target, 1)) : 1;
    var name = isExtra ? k.name : line.counterpart;
    return {
      id: k.id, lineId: line.id, lineName: line.name, customerId: c.id, customerName: c.name,
      name: name, text: kpiText(name, k), target: k.target, unit: k.unit, start: k.start, end: k.end,
      source: k.source, owner: k.owner, money: E.amount(line, 'today'),
      plan: (k.plan || rep(0)).slice(), done: (k.done || rep(null)).slice(), proof: k.proof || 'None',
      tracked: k.tracked === true, checked: k.checked === true, due: k.due === true, paidFull: k.paidFull === true,
      actual: actual, isExtra: isExtra, elapsed: E.elapsed(k.start, k.end), expected: expected, status: status, kind: kind, delivered: delivered
    };
  }
  // pay against execution for one line (conditional today) and its KPIs
  function linePay(line, ks) {
    var amount = E.amount(line, 'today');
    var paid = finite(line.paid) ? line.paid : null;
    var paidShare = paid !== null && amount > 0 ? paid / amount : 0;
    var delivered = ks.reduce(function (m, k) { return Math.min(m, k.delivered); }, 1);
    var ahead = paidShare > delivered + 1e-9;
    return { lineId: line.id, lineName: line.name, amount: amount, paid: paid, paidShare: paidShare, delivered: delivered,
      ahead: ahead, excess: ahead ? paid - amount * delivered : 0 };
  }
  E.kpis = function (c) {
    var out = [];
    if (!c) return out;
    c.lines.forEach(function (line) {
      if (E.classOf(line, 'today') !== 'conditional') return;
      var ks = [mkKpi(c, line, line, false)];
      (line.kpis || []).forEach(function (k) { if (E.kpiComplete(k)) ks.push(mkKpi(c, line, k, true)); });
      var p = linePay(line, ks);
      ks.forEach(function (k) {
        k.paid = p.paid; k.paidShare = p.paidShare; k.lineDelivered = p.delivered; k.ahead = p.ahead; k.excess = p.excess;
        out.push(k);
      });
    });
    return out;
  };
  // one entry per line that is conditional today
  E.paidAhead = function (c) {
    var out = [], seen = {};
    E.kpis(c).forEach(function (k) {
      if (seen[k.lineId]) return;
      seen[k.lineId] = 1;
      out.push({ lineId: k.lineId, lineName: k.lineName, amount: k.money, paid: k.paid, paidShare: k.paidShare,
        delivered: k.lineDelivered, ahead: k.ahead, excess: k.excess });
    });
    return out;
  };
  E.paidAheadTotal = function (customers) {
    var list = Array.isArray(customers) ? customers : customers ? [customers] : S().customers;
    return list.reduce(function (s, c) { return s + E.paidAhead(c).reduce(function (t, p) { return t + p.excess; }, 0); }, 0);
  };
  E.allKpis = function () { var out = []; S().customers.forEach(function (c) { out = out.concat(E.kpis(c)); }); return out; };
  E.moneyAtStake = function (kpis) {
    var seen = {}, sum = 0;
    (kpis || []).forEach(function (k) { var key = k.customerId + '|' + k.lineId; if (!seen[key]) { seen[key] = 1; sum += k.money; } });
    return sum;
  };
  E.moneyByStatus = function (kpis) {
    var out = { executed: 0, partly: 0, notexecuted: 0 };
    Object.keys(out).forEach(function (st) { out[st] = E.moneyAtStake((kpis || []).filter(function (k) { return k.status === st; })); });
    return out;
  };
  // the key figures of one customer or several (SPEC5 3.5), all from today's agreement except linkedProposed
  E.keyFigures = function (x) {
    var list = Array.isArray(x) ? x : x ? [x] : S().customers;
    var lines = [], ks = [];
    list.forEach(function (c) { lines = lines.concat(c.lines); ks = ks.concat(E.kpis(c)); });
    var s = E.summary(lines, 'today'), sales = E.sales(list, 'today');
    var cp = { lines: 0, smart: 0, aligned: 0, smartNotAlignedAmount: 0 };
    lines.forEach(function (l) {
      if (E.type(l, 'today') !== 'conditional') return;
      var sm = E.smart(l, 'today').score === 4, al = E.aligned(l, 'today');
      cp.lines += 1;
      if (sm) cp.smart += 1;
      if (al) cp.aligned += 1;
      if (sm && !al) cp.smartNotAlignedAmount += E.amount(l, 'today');
    });
    var kp = { n: ks.length, executed: 0, partly: 0, notexecuted: 0, atStake: E.moneyAtStake(ks) };
    ks.forEach(function (k) { kp[k.status] += 1; });
    return {
      total: s.total, pctOfGross: sales.gross > 0 ? s.total / sales.gross : 0, sales: sales,
      classes: { unconditional: { amount: s.unconditional, share: s.shareUnconditional }, partial: { amount: s.partial, share: s.sharePartial },
        conditional: { amount: s.conditional, share: s.shareConditional } },
      counterparts: cp, kpis: kp, paidAhead: E.paidAheadTotal(list),
      linkedToday: E.linkedShare(lines, 'today'), linkedProposed: E.linkedShare(lines, 'proposed')
    };
  };
  // investment growth against net invoice sales growth (the first net today against revenueLY, net invoice sales last year)
  E.growth = function (c) {
    var total = E.summary(c.lines, 'today').total, net = E.sales(c, 'today').firstNet;
    var inv = c.investmentLY > 0 ? total / c.investmentLY - 1 : 0;
    var rev = c.revenueLY > 0 ? net / c.revenueLY - 1 : 0;
    var excess = c.revenueLY > 0 ? total - c.investmentLY * (net / c.revenueLY) : 0;
    return { investment: inv, revenue: rev, faster: inv > rev, excess: excess };
  };
  E.opportunitySize = function (c) {
    var total = 0, good = 0;
    c.lines.forEach(function (l) {
      var a = E.amount(l, 'today');
      total += a;
      if (E.classOf(l, 'today') === 'conditional' && E.priorityId(l, 'today')) good += a;
    });
    return total - good;
  };
  E.suggestedMustWin = function (c) {
    var best = null;
    E.spendByPriority(c.lines, 'today').forEach(function (e) {
      if (!e.priorityId) return;
      if (!best || e.amount < best.amount - 1e-9) best = e;
    });
    return best ? best.priorityId : null;
  };
  E.mustWin = function (c) { return (c.mustWinId && E.priority(c.mustWinId)) ? c.mustWinId : E.suggestedMustWin(c); };
  var MISSING_WORD = { counterpart: 'a counterpart', target: 'a target', date: 'a date', owner: 'an owner' };
  E.opportunities = function (c) {
    var items = [], lines = c.lines;
    lines.filter(function (l) { return E.classOf(l, 'today') === 'unconditional' && E.amount(l, 'today') > 0; })
      .sort(function (a, b) { return E.amount(b, 'today') - E.amount(a, 'today'); })
      .slice(0, 2)
      .forEach(function (l) {
        var a = E.amount(l, 'today');
        items.push({ kind: 'unconditional', money: a, lineIds: [l.id], title: fmt.money(a) + ' unconditional', detail: l.name + ' has no counterpart.' });
      });
    lines.filter(function (l) { return E.classOf(l, 'today') === 'partial'; }).forEach(function (l) {
      var a = E.amount(l, 'today');
      var miss = E.missing(l).map(function (w) { return MISSING_WORD[w]; });
      items.push({ kind: 'partial', money: a, lineIds: [l.id], title: fmt.money(a) + ' partially conditional', detail: l.name + ': add ' + listJoin(miss) + '.' });
    });
    var np = lines.filter(function (l) { return E.classOf(l, 'today') === 'conditional' && !E.priorityId(l, 'today'); });
    var npSum = np.reduce(function (s, l) { return s + E.amount(l, 'today'); }, 0);
    if (npSum > 0) {
      items.push({ kind: 'nopriority', money: npSum, lineIds: np.map(function (l) { return l.id; }),
        title: fmt.money(npSum) + ' conditional, no priority',
        detail: listJoin(np.map(function (l) { return l.name; })) + (np.length > 1 ? ' serve' : ' serves') + ' none of the priorities.' });
    }
    var g = E.growth(c);
    if (g.faster) {
      items.push({ kind: 'growth', money: g.excess, lineIds: [], title: fmt.money(g.excess) + ' above revenue growth',
        detail: 'Investment grew ' + fmt.pct1(g.investment) + ', revenue ' + fmt.pct1(g.revenue) + '.' });
    }
    E.spendByPriority(lines, 'today').forEach(function (e) {
      if (e.priorityId && e.amount <= 0) {
        var p = E.priority(e.priorityId);
        items.push({ kind: 'gap', money: null, lineIds: [], priorityId: e.priorityId, title: 'No spend on ' + p.name, detail: p.winning || '' });
      }
    });
    items.sort(function (a, b) {
      if (a.money === null && b.money === null) return 0;
      if (a.money === null) return 1;
      if (b.money === null) return -1;
      return b.money - a.money;
    });
    return items.slice(0, 6);
  };
  TTS.engine = E;

  // the proposed action (SPEC5 3.2): its effect is applied once, then the proposed fields stay editable
  TTS.applyAction = function (line, action) {
    if (!line || ACTIONS.indexOf(action) < 0) return line;
    line.action = action;
    if (action === 'Keep') {
      line.propAmount = finite(line.amount) ? line.amount : 0; line.propType = E.type(line, 'today'); line.propPriorityId = E.priorityId(line, 'today');
    } else if (action === 'Make conditional' || action === 'Strengthen the counterpart') {
      line.propType = 'conditional';
    } else if (action === 'Fold into the list price' || action === 'Remove') {
      line.propAmount = 0;
    }
    return line;
  };

  // SPEC5b A: the one place that recalculates the amounts of the variable lines (amount = rate × base, one decimal): the
  // on-invoice lines first (base: gross revenue), then the first net, then every other line (base: the first net). A variable
  // line with no rate yet takes its rate from its amount. With pack groups, promoSpend = the sum of their promo.
  // Called on load, on import and by TTS.commit(); no argument = every customer. Returns the customer (or the state).
  TTS.recalc = function (c) {
    if (!c) { ((S() && S().customers) || []).forEach(function (x) { TTS.recalc(x); }); return S(); }
    var lines = c.lines || [], gross = finite(c.listSales) ? c.listSales : 0;
    function apply(l, base) {
      if (l.basis !== 'Variable') return;
      if (finite(l.rate)) l.amount = r1(l.rate / 100 * base);
      else l.rate = base > 0 ? r2((finite(l.amount) ? l.amount : 0) / base * 100) : 0;
    }
    lines.forEach(function (l) { if (l.paidHow === 'On invoice') apply(l, gross); });
    var firstNet = gross - lines.reduce(function (s, l) { return s + (l.paidHow === 'On invoice' && finite(l.amount) ? l.amount : 0); }, 0);
    lines.forEach(function (l) { if (l.paidHow !== 'On invoice') apply(l, firstNet); });
    if (Array.isArray(c.packs) && c.packs.length) {
      c.promoSpend = r1(c.packs.reduce(function (s, p) { return s + (finite(p.promo) ? p.promo : 0); }, 0));
    }
    return c;
  };
  // switches a line between 'Fixed' and 'Variable' (SPEC5b A): to Variable the rate becomes the line's current rate (two
  // decimals) so the amount stays; to Fixed the amount stays. Returns the line. (The caller commits.)
  TTS.setBasis = function (c, line, basis) {
    if (!line || BASES.indexOf(basis) < 0 || line.basis === basis) return line;
    if (basis === 'Variable') { line.rate = E.rateOf(line, c); if (!finite(line.rate)) line.rate = 0; }
    else line.rate = null;
    line.basis = basis;
    return line;
  };
  // the customer's promotional spend (SPEC5b A). With pack groups the typed total rescales their promo in proportion (by share
  // when they are all 0), one decimal each, the rounding rest on the largest, so the sum stays the typed total. (The caller commits.)
  TTS.setPromoSpend = function (c, value) {
    var v = finite(value) ? Math.max(0, r1(value)) : 0;
    c.promoSpend = v;
    var ps = Array.isArray(c.packs) ? c.packs : [];
    if (!ps.length) return c;
    var sum = ps.reduce(function (s, p) { return s + (finite(p.promo) ? p.promo : 0); }, 0);
    var shares = ps.reduce(function (s, p) { return s + (finite(p.share) ? p.share : 0); }, 0);
    var w = ps.map(function (p) { return sum > 0 ? (finite(p.promo) ? p.promo : 0) / sum : shares > 0 ? (finite(p.share) ? p.share : 0) / shares : 1 / ps.length; });
    ps.forEach(function (p, i) { p.promo = r1(v * w[i]); });
    var big = 0;
    w.forEach(function (x, i) { if (x > w[big]) big = i; });
    ps[big].promo = r1(ps[big].promo + v - ps.reduce(function (s, p) { return s + p.promo; }, 0));
    return c;
  };

  // ---------------------------------------------------------------- state helpers
  TTS.uid = uid;
  TTS.listJoin = listJoin;
  TTS.dayNum = dayNum;
  TTS.customer = function (id) {
    var cs = (S() && S().customers) || [];
    for (var i = 0; i < cs.length; i++) if (cs[i].id === id) return cs[i];
    return null;
  };
  TTS.scopeLines = function (scope) {
    if (scope && scope !== 'all') { var c = TTS.customer(scope); return c ? c.lines : []; }
    var out = [];
    S().customers.forEach(function (c) { out = out.concat(c.lines); });
    return out;
  };
  function findLine(c, id) { if (!c) return null; for (var i = 0; i < c.lines.length; i++) if (c.lines[i].id === id) return c.lines[i]; return null; }
  function findKpi(line, id) { if (!line) return null; var ks = line.kpis || []; for (var i = 0; i < ks.length; i++) if (ks[i].id === id) return ks[i]; return null; }
  function isUnchanged() {
    var ex = exampleState(), st = S();
    return JSON.stringify([st.priorities, st.customers, st.year, st.demoDate]) === JSON.stringify([ex.priorities, ex.customers, ex.year, ex.demoDate]);
  }

  // ---------------------------------------------------------------- persistence
  function valid(st) {
    return st && st.version === 5 && Array.isArray(st.customers) && Array.isArray(st.priorities) && finite(st.year) && typeof st.demoDate === 'string';
  }
  // fills every missing customer and line field with its default, so an imported or hand-edited state never breaks
  function normCustomer(c, y, demo) {
    c.lines = Array.isArray(c.lines) ? c.lines : []; c.scriptEdits = c.scriptEdits || {};
    c.lines.forEach(function (l) { normLine(l, y, demo); });
    if (!finite(c.listSales)) {
      var onInv = c.lines.reduce(function (s, l) { return s + (l.paidHow === 'On invoice' && finite(l.amount) ? l.amount : 0); }, 0);
      c.listSales = finite(c.revenueTY) ? c.revenueTY + onInv : 0;
    }
    delete c.revenueTY;
    if (!finite(c.revenueLY)) c.revenueLY = 0;
    if (!finite(c.investmentLY)) c.investmentLY = 0;
    // SPEC5b A: promotional spend (0 by default) and the optional pack groups
    if (!finite(c.promoSpend)) c.promoSpend = 0;
    if (Array.isArray(c.packs)) {
      c.packs = c.packs.filter(function (p) { return p && typeof p === 'object'; }).map(function (p) {
        return { name: typeof p.name === 'string' ? p.name : '', share: finite(p.share) ? p.share : 0, promo: finite(p.promo) ? p.promo : 0 };
      });
      if (!c.packs.length) delete c.packs;
    } else delete c.packs;
    TTS.recalc(c);
    return c;
  }
  TTS.load = function () {
    var st = null;
    try {
      if (global.localStorage) global.localStorage.removeItem('tts.state.v4');   // the old example data (SPEC10 5)
      var raw = global.localStorage && global.localStorage.getItem(STORE_KEY);
      if (raw) st = JSON.parse(raw);
    } catch (e) { st = null; }
    if (!valid(st)) st = exampleState();
    st.ui = st.ui || {};
    if (st.ui.agreementMode !== 'proposed') st.ui.agreementMode = 'today';
    st.customers.forEach(function (c) { normCustomer(c, st.year, st.demoDate); });
    TTS.state = st;
    return st;
  };
  TTS.save = function () {
    try { if (global.localStorage) global.localStorage.setItem(STORE_KEY, JSON.stringify(TTS.state)); } catch (e) { /* session only */ }
  };
  TTS.resetToExample = function () {
    var ui = (TTS.state && TTS.state.ui) || {};
    TTS.state = exampleState();
    TTS.state.ui.agreementMode = ui.agreementMode === 'proposed' ? 'proposed' : 'today';
    TTS.lastImport = null;
    TTS.details = {};
    TTS.save();
  };

  // ---------------------------------------------------------------- UI helpers
  function esc(s) {
    if (s === null || s === undefined) return '';
    return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  }
  var ICONS = {
    download: '<path d="M12 3v12"/><path d="m7 10 5 5 5-5"/><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>',
    upload: '<path d="M12 15V3"/><path d="m7 8 5-5 5 5"/><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>',
    plus: '<path d="M12 5v14"/><path d="M5 12h14"/>',
    trash: '<path d="M4 7h16"/><path d="M10 11v6"/><path d="M14 11v6"/><path d="M6 7l1 12a2 2 0 0 0 2 2h6a2 2 0 0 0 2-2l1-12"/><path d="M9 7V4h6v3"/>',
    print: '<path d="M7 9V3h10v6"/><rect x="3" y="9" width="18" height="8" rx="2"/><path d="M7 14h10v7H7z"/>',
    copy: '<rect x="9" y="9" width="12" height="12" rx="2"/><path d="M5 15H4a1 1 0 0 1-1-1V4a1 1 0 0 1 1-1h10a1 1 0 0 1 1 1v1"/>',
    reset: '<path d="M3 12a9 9 0 1 0 3-6.7L3 8"/><path d="M3 3v5h5"/>',
    'caret-down': '<path d="m6 9 6 6 6-6"/>',
    check: '<path d="m5 12 5 5 9-10"/>',
    'arrow-right': '<path d="M5 12h14"/><path d="m13 6 6 6-6 6"/>',
    calendar: '<rect x="3" y="5" width="18" height="16" rx="2"/><path d="M3 10h18"/><path d="M8 3v4"/><path d="M16 3v4"/>',
    alert: '<path d="M12 3 2 20h20L12 3z"/><path d="M12 10v4"/><path d="M12 17h.01"/>',
    file: '<path d="M14 3H7a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V8z"/><path d="M14 3v5h5"/>',
    cross: '<path d="M7 7l10 10"/><path d="M17 7 7 17"/>'
  };
  function icon(name, size) {
    var s = size || 18;
    return '<svg class="ic" width="' + s + '" height="' + s + '" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" ' +
      'stroke-linecap="round" stroke-linejoin="round" aria-hidden="true" focusable="false">' + (ICONS[name] || '') + '</svg>';
  }
  function chip(kind, label) { return '<span class="chip ' + esc(kind) + '">' + esc(label) + '</span>'; }
  var LINE_CHIP = { unconditional: ['grey', 'Unconditional'], conditional: ['green', 'Conditional'], partial: ['yellow', 'Partially conditional'] };
  var STATUS_CHIP = { executed: ['green', 'Fully executed'], partly: ['yellow', 'Partially executed'], notexecuted: ['outline', 'Not executed'] };
  function lineChip(status) { var c = LINE_CHIP[status]; return chip(c[0], c[1]); }
  function statusChip(status) { var c = STATUS_CHIP[status]; return chip(c[0], c[1]); }

  // items: [{ value, pre, unit, label, sub, good }]
  function statBand(items, rightHtml) {
    var h = items.map(function (it) {
      return '<div class="stat' + (it.good ? ' good' : '') + (it.cls ? ' ' + it.cls : '') + '"><div class="v">' +
        (it.pre ? '<span class="u pre">' + esc(it.pre) + '</span>' : '') + esc(it.value) +
        (it.unit ? '<span class="u">' + esc(it.unit) + '</span>' : '') + '</div>' +
        '<div class="l">' + esc(it.label) + (it.sub ? ' <span class="t"><span class="tsep">· </span>' + esc(it.sub) + '</span>' : '') + '</div></div>';
    }).join('');
    return '<div class="stats">' + h + (rightHtml || '') + '</div>';
  }
  function splitBar(s, opts) {
    var o = opts || {};
    var segs = [['c', s.shareConditional], ['n', s.sharePartial], ['u', s.shareUnconditional]].filter(function (x) { return x[1] > 0; })
      .map(function (x) { return '<i class="' + x[0] + '" style="flex-grow:' + x[1].toFixed(4) + '"></i>'; }).join('');
    var bar = '<div class="bar' + (segs ? '' : ' empty') + '">' + segs + '</div>';
    if (o.legend === false) return '<div class="split bare">' + bar + '</div>';
    return '<div class="split">' + bar + '<div class="legend">' +
      '<span><i class="dot c"></i>Conditional with a number <b>' + fmt.pct(s.shareConditional) + '</b></span>' +
      '<span><i class="dot n"></i>Partially conditional <b>' + fmt.pct(s.sharePartial) + '</b></span>' +
      '<span><i class="dot u"></i>Unconditional <b>' + fmt.pct(s.shareUnconditional) + '</b></span></div></div>';
  }
  // one stacked bar of spend by bucket (SPEC4 4.2 and 5); width = this total against `max` (the same scale for both bars)
  function bucketBar(entries, label, max) {
    var total = entries.reduce(function (s, e) { return s + Math.max(e.amount, 0); }, 0);
    var shown = entries.filter(function (e) { return e.amount > 1e-9; });
    var segs = shown.map(function (e) {
      return '<i class="bk-' + e.key + '" style="flex-grow:' + (e.amount / total).toFixed(4) + '" title="' + esc(e.name + ': ' + fmt.money(e.amount) + ', ' + fmt.pct(e.share)) + '"></i>';
    }).join('');
    var aria = label + ': ' + (shown.length ? shown.map(function (e) { return e.name + ' ' + fmt.m(e.amount); }).join(', ') : 'no spend');
    var w = max > 0 ? Math.max(total / max * 100, 0) : 100;
    return '<div class="bk-bar' + (segs ? '' : ' empty') + '" role="img" aria-label="' + esc(aria) + '"' + (w < 99.995 ? ' style="width:' + w.toFixed(2) + '%"' : '') + '>' + segs + '</div>';
  }
  // the two bars, Today above Proposed, and a legend with each bucket's SEK m today → proposed
  function bucketChart(lines) {
    var t = E.spendByBucket(lines, 'today'), p = E.spendByBucket(lines, 'proposed');
    var sum = function (es) { return es.reduce(function (s, e) { return s + Math.max(e.amount, 0); }, 0); };
    var max = Math.max(sum(t), sum(p));
    var bars = '<div class="bk-bars"><span class="bk-lab">Today</span>' + bucketBar(t, 'Today', max) +
      '<span class="bk-lab">Proposed</span>' + bucketBar(p, 'Proposed', max) + '</div>';
    var items = [];
    t.forEach(function (e, i) {
      var a = r1(e.amount), b = r1(p[i].amount);
      if (e.key === 'nobucket' && !(a > 0 || b > 0)) return;
      var better = e.key === 'unconditional' ? b < a : b > a;
      var zero = a === 0 && b === 0;
      items.push('<span class="bk-name' + (zero ? ' zero' : '') + '"><i class="dot bk-' + e.key + '"></i>' + esc(e.name) + '</span>' +
        '<span class="bk-num' + (zero ? ' zero' : '') + '">' + fmt.m(a) + '</span><span class="bk-to' + (zero ? ' zero' : '') + '">' + icon('arrow-right', 12) + '</span>' +
        '<span class="bk-num' + (zero ? ' zero' : better ? ' better' : '') + '">' + fmt.m(b) + '</span>');
    });
    var rows = Math.ceil(items.length / 2), cells = '';
    for (var r = 0; r < rows; r++) {
      cells += items[r] + (items[r + rows] ? items[r + rows].replace('<span class="bk-name', '<span class="g2 bk-name') : '');
    }
    return '<div class="bk-chart">' + bars + '<div class="bk-legend">' + cells + '</div></div>';
  }
  // the bucket bar of one table row: full cell width, shares of that row's total
  function bucketRowBar(lines, mode) {
    return bucketBar(E.spendByBucket(lines, mode), mode === 'proposed' ? 'Proposed' : 'Today', 0);
  }
  function priorityOptions(selectedId) {
    var h = '<option value=""' + (selectedId ? '' : ' selected') + '>No priority</option>';
    S().priorities.forEach(function (p) {
      h += '<option value="' + esc(p.id) + '"' + (p.id === selectedId ? ' selected' : '') + '>' + esc(p.name) + '</option>';
    });
    return h;
  }
  function emptyState(o) {
    var b = o.button;
    var btn = '';
    if (b) {
      btn = b.href ? '<a class="btn primary" href="' + esc(b.href) + '">' + esc(b.label) + '</a>'
        : '<button type="button" class="btn primary" data-action="' + esc(b.action) + '">' + esc(b.label) + '</button>';
    }
    return '<div class="empty">' + '<span class="empty-ic">' + icon(o.icon || 'file', 40) + '</span>' +
      '<h2>' + esc(o.title) + '</h2><p>' + esc(o.text) + '</p>' + btn + '</div>';
  }
  var toastTimer = null;
  function toast(msg) {
    if (!HAS_DOM) return;
    var t = document.getElementById('tts-toast');
    if (!t) return;
    t.textContent = msg;
    t.classList.add('show');
    clearTimeout(toastTimer);
    toastTimer = setTimeout(function () { t.classList.remove('show'); }, 2500);
  }
  function confirmDlg(o) {
    return new Promise(function (resolve) {
      var d = HAS_DOM && document.getElementById('tts-dialog');
      if (!d || typeof d.showModal !== 'function') { resolve(false); return; }
      d.innerHTML = '<form method="dialog"><h2>' + esc(o.title) + '</h2><p>' + esc(o.text || '') + '</p>' +
        '<div class="dlg-actions"><button type="button" class="btn secondary" value="cancel">Cancel</button>' +
        '<button type="button" class="btn primary" value="ok">' + esc(o.okLabel || 'OK') + '</button></div></form>';
      var done = false;
      function finish(v) {
        if (done) return; done = true;
        if (d.open) d.close();
        resolve(v);
      }
      d.querySelectorAll('button').forEach(function (b) { b.addEventListener('click', function () { finish(b.value === 'ok'); }); });
      d.onclose = function () { finish(false); };
      d.oncancel = function () { finish(false); };
      d.showModal();
      var cancel = d.querySelector('button[value="cancel"]');
      if (cancel) cancel.focus();
    });
  }
  TTS.ui = {
    esc: esc, icon: icon, chip: chip, lineChip: lineChip, statusChip: statusChip, statBand: statBand, splitBar: splitBar,
    bucketChart: bucketChart, bucketRowBar: bucketRowBar,
    toast: toast, confirm: confirmDlg, priorityOptions: priorityOptions, emptyState: emptyState, bottomRow: function (left) { return bottomRow(left); },
    money: function (n) { return fmt.money(n); }
  };

  // ---------------------------------------------------------------- Excel
  var LINE_HEAD = ['Investment line', 'Amount (SEK m)', 'Paid how', 'Bucket', 'Fixed or variable', 'Rate (%)', 'Type', 'Priority', 'Counterpart', 'Target', 'Unit',
    'From', 'To', 'Data source', 'Owner', 'Actual to date', 'Paid to date (SEK m)', 'Proposed amount (SEK m)', 'Proposed type', 'Proposed priority', 'KPI in full'];
  var LINE_W = [26, 15, 14, 18, 17, 10, 15, 24, 34, 9, 14, 12, 12, 26, 18, 14, 19, 22, 15, 24, 60];
  var KPI_HEAD = ['Investment line', 'KPI', 'Target', 'Unit', 'From', 'To', 'Data source', 'Owner', 'Actual to date', 'KPI in full'];
  var KPI_W = [26, 34, 9, 14, 12, 12, 26, 18, 14, 60];
  function typeLabel(t) { return t === 'conditional' ? 'Conditional' : 'Unconditional'; }
  function pName(id) { var p = E.priority(id); return p ? p.name : ''; }
  function xDate(iso) { if (!iso) return null; var p = iso.split('-'); return new Date(+p[0], +p[1] - 1, +p[2]); }
  function xNum(n) { return finite(n) ? n : null; }
  function mkSheet(XLSX, rows, widths) {
    var ws = XLSX.utils.aoa_to_sheet(rows, { cellDates: true, dateNF: 'yyyy-mm-dd' });
    ws['!cols'] = widths.map(function (w) { return { wch: w }; });
    return ws;
  }
  function customerRows(c) {
    return [['Field', 'Value'], ['Customer', c ? c.name : ''], ['Channel', c ? c.channel : ''],
      ['Revenue last year (SEK m)', c ? xNum(c.revenueLY) : null], ['Sales at list price (SEK m)', c ? xNum(c.listSales) : null],
      ['Investment last year (SEK m)', c ? xNum(c.investmentLY) : null], ['Promotional spend (SEK m)', c ? xNum(c.promoSpend) : null]];
  }
  var excel = {};
  excel.templateWorkbook = function (XLSX, state) {
    var prev = TTS.state; TTS.state = state || prev;
    var wb = XLSX.utils.book_new();
    var names = TTS.state.priorities.map(function (p) { return p.name; }).join('; ');
    XLSX.utils.book_append_sheet(wb, mkSheet(XLSX, [
      ['Read me'],
      ['1. One file per customer. Fill in the sheets Customer, Lines and More KPIs, then import the file in the Trade terms simulator.'],
      ['2. Customer: the customer name, the channel (Home or Away from Home), revenue last year (net invoice sales), sales at list price this year and investment last year, all in SEK m.'],
      ['3. Lines: one row per investment line. Type is Unconditional or Conditional. Dates as Excel dates or YYYY-MM-DD.'],
      ['4. A conditional line needs a counterpart, a target number, a To date and an owner. Without them it shows as "Partially conditional".'],
      ['5. More KPIs (optional): extra KPIs for a line, matched on the investment line name. Blank proposed fields keep today\'s values.'],
      ['6. Priority is blank or one of: ' + names + '.'],
      ['7. Paid how is blank, On invoice, Off invoice or Fixed amount. Paid to date is the money paid so far on the line, in SEK m.'],
      ['8. Bucket is what the money of a conditional line is for: blank, Listings, Shelf space, In-store marketing, Promotions or Efficiency.'],
      ['9. Fixed or variable is Fixed (you give the amount) or Variable (you give the rate). Rate (%) is the rate of a variable line: % of sales at list price for an on-invoice line, % of net invoice sales for any other line. The amount is then calculated and may be left blank.'],
      ['10. Promotional spend (Customer sheet) is your estimate of the money spent on price discounts in promotions this year, in SEK m. It is not a contract point. Blank means 0.'],
      ['11. KPI in full (optional, last column of Lines and More KPIs): the KPI in specific words, for example "10 leaflet weeks, each with a floor display in all large stores". Blank means the tool writes it from the counterpart, the target and the unit.']
    ], [130]), 'Read me');
    XLSX.utils.book_append_sheet(wb, mkSheet(XLSX, customerRows(null), [30, 30]), 'Customer');
    XLSX.utils.book_append_sheet(wb, mkSheet(XLSX, [LINE_HEAD], LINE_W), 'Lines');
    XLSX.utils.book_append_sheet(wb, mkSheet(XLSX, [KPI_HEAD], KPI_W), 'More KPIs');
    TTS.state = prev;
    return wb;
  };
  excel.customerWorkbook = function (XLSX, state, c) {
    var prev = TTS.state; TTS.state = state || prev;
    var wb = XLSX.utils.book_new();
    var lines = [LINE_HEAD], kpis = [KPI_HEAD];
    c.lines.forEach(function (l) {
      var variable = l.basis === 'Variable';
      lines.push([l.name, xNum(l.amount), l.paidHow || '', l.bucket || '', variable ? 'Variable' : 'Fixed', variable ? xNum(l.rate) : null,
        typeLabel(l.type), pName(l.priorityId), l.counterpart || '', xNum(l.target), l.unit || '',
        xDate(l.start), xDate(l.end), l.source || '', l.owner || '', xNum(E.actualOf(l)), xNum(l.paid), xNum(l.propAmount), typeLabel(l.propType), pName(l.propPriorityId), l.kpiFull || '']);
      (l.kpis || []).forEach(function (k) {
        kpis.push([l.name, k.name || '', xNum(k.target), k.unit || '', xDate(k.start), xDate(k.end), k.source || '', k.owner || '', xNum(E.actualOf(k)), k.kpiFull || '']);
      });
    });
    XLSX.utils.book_append_sheet(wb, mkSheet(XLSX, customerRows(c), [30, 30]), 'Customer');
    XLSX.utils.book_append_sheet(wb, mkSheet(XLSX, lines, LINE_W), 'Lines');
    XLSX.utils.book_append_sheet(wb, mkSheet(XLSX, kpis, KPI_W), 'More KPIs');
    TTS.state = prev;
    return wb;
  };
  excel.exportName = function (state, scope) {
    var st = state || TTS.state;
    if (!scope || scope === 'all') return 'Trade terms - all customers.xlsx';
    var c = null;
    st.customers.forEach(function (x) { if (x.id === scope) c = x; });
    return 'Trade terms - ' + (c ? c.name : scope).replace(/[\\/:*?"<>|]/g, '') + '.xlsx';
  };
  excel.exportWorkbook = function (XLSX, state, scope) {
    var prev = TTS.state; TTS.state = state || prev;
    var cs = (!scope || scope === 'all') ? TTS.state.customers : TTS.state.customers.filter(function (c) { return c.id === scope; });
    var ag = [['Customer', 'Channel', 'Investment line', 'Amount (SEK m)', 'Paid how', 'Bucket', 'Type', 'Status', 'Priority', 'Counterpart', 'Target', 'Unit', 'From', 'To',
      'Data source', 'Owner', 'Paid to date (SEK m)', 'Proposed amount (SEK m)', 'Proposed type', 'Proposed priority']];
    var kp = [['Customer', 'Investment line', 'KPI', 'Target', 'Unit', 'From', 'To', 'Data source', 'Owner', 'SEK m at stake', 'Actual to date',
      'Expected by now', 'Status', 'Paid to date (SEK m)', 'Paid ahead of delivery']];
    cs.forEach(function (c) {
      c.lines.forEach(function (l) {
        ag.push([c.name, c.channel, l.name, xNum(l.amount), l.paidHow || '', l.bucket || '', typeLabel(l.type), E.CLASS_LABEL[E.classOf(l, 'today')], pName(l.priorityId), l.counterpart || '',
          xNum(l.target), l.unit || '', xDate(l.start), xDate(l.end), l.source || '', l.owner || '', xNum(l.paid), xNum(l.propAmount), typeLabel(l.propType), pName(l.propPriorityId)]);
      });
      E.kpis(c).forEach(function (k) {
        kp.push([c.name, k.lineName, k.text, xNum(k.target), k.unit || '', xDate(k.start), xDate(k.end), k.source || '', k.owner || '',
          r1(k.money), k.actual, k.actual === null || k.expected === null ? null : r1(k.expected), E.statusLabel(k.status),
          xNum(k.paid), k.ahead ? 'Yes' : null]);
      });
    });
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, mkSheet(XLSX, ag, [16, 16, 26, 15, 14, 18, 15, 14, 24, 34, 9, 14, 12, 12, 26, 18, 19, 22, 15, 24]), 'Agreement');
    XLSX.utils.book_append_sheet(wb, mkSheet(XLSX, kp, [16, 26, 34, 9, 14, 12, 12, 26, 18, 14, 14, 16, 14, 19, 22]), 'KPIs');
    TTS.state = prev;
    return wb;
  };

  function norm(s) { return String(s === null || s === undefined ? '' : s).trim().toLowerCase().replace(/\s+/g, ' '); }
  function sheetByName(wb, name) {
    var n = norm(name);
    for (var i = 0; i < wb.SheetNames.length; i++) if (norm(wb.SheetNames[i]) === n) return wb.Sheets[wb.SheetNames[i]];
    return null;
  }
  function sheetRows(XLSX, ws) {
    if (!ws || !ws['!ref']) return { rows: [], base: 0 };
    var range = XLSX.utils.decode_range(ws['!ref']);
    var rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: true });
    return { rows: rows, base: range.s.r };
  }
  function cellText(v) {
    if (v === null || v === undefined) return '';
    if (v instanceof Date) return isNaN(v.getTime()) ? '' : isoOf(v);
    return String(v).trim();
  }
  function isoOf(d) {
    var t = new Date(d.getTime() + 43200000); // half a day: absorbs time-zone and rounding offsets
    function p2(n) { return (n < 10 ? '0' : '') + n; }
    return t.getFullYear() + '-' + p2(t.getMonth() + 1) + '-' + p2(t.getDate());
  }
  function readDate(XLSX, v) {
    if (v === null || v === undefined || (typeof v === 'string' && v.trim() === '')) return { ok: true, value: null };
    if (v instanceof Date) return isNaN(v.getTime()) ? { ok: false } : { ok: true, value: isoOf(v) };
    if (typeof v === 'number' && XLSX.SSF && XLSX.SSF.parse_date_code) {
      var dc = XLSX.SSF.parse_date_code(v);
      if (dc && dc.y > 1900) { var d = new Date(dc.y, dc.m - 1, dc.d); return { ok: true, value: isoOf(d) }; }
      return { ok: false };
    }
    var s = String(v).trim();
    var m = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(s);
    if (m) {
      var dt = new Date(+m[1], +m[2] - 1, +m[3]);
      if (dt.getMonth() === +m[2] - 1 && dt.getDate() === +m[3]) return { ok: true, value: isoOf(dt) };
    }
    return { ok: false };
  }

  excel.parseCustomerWorkbook = function (XLSX, wb, state, fileName) {
    var prev = TTS.state; TTS.state = state || prev;
    try { return parseWb(XLSX, wb, fileName); } finally { TTS.state = prev; }
  };
  function parseWb(XLSX, wb, fileName) {
    var messages = [];
    function fail(text) {
      messages.push({ level: 'error', row: null, sheet: null, text: text });
      return { ok: false, customer: null, linesRead: 0, kpisRead: 0, messages: messages };
    }
    var linesWs = sheetByName(wb, 'Lines');
    if (!linesWs) return fail('This file has no sheet called "Lines". Download the template to see the standard layout.');

    // Customer sheet
    var cws = sheetByName(wb, 'Customer'), fields = {};
    if (cws) {
      sheetRows(XLSX, cws).rows.forEach(function (r) { if (r && r.length) fields[norm(r[0])] = r[1]; });
    }
    var name = cellText(fields.customer);
    if (!name) return fail('The "Customer" sheet has no customer name.');

    var lr = sheetRows(XLSX, linesWs);
    var head = (lr.rows[0] || []).map(norm);
    function col(h) { return head.indexOf(norm(h)); }
    var REQ = ['Investment line', 'Amount (SEK m)', 'Type'];
    for (var q = 0; q < REQ.length; q++) {
      if (col(REQ[q]) < 0) return fail('The "Lines" sheet has no column called "' + REQ[q] + '".');
    }
    var ci = {};
    LINE_HEAD.forEach(function (h) { ci[h] = col(h); });
    function get(r, h) { var i = ci[h]; return i >= 0 ? r[i] : null; }
    function findPriority(v) {
      var n = norm(v);
      var ps = TTS.state.priorities;
      for (var i = 0; i < ps.length; i++) if (norm(ps[i].name) === n) return ps[i].id;
      return undefined;
    }
    function parseType(v) { var n = norm(v); return n === 'conditional' ? 'conditional' : n === 'unconditional' ? 'unconditional' : null; }

    var lines = [];
    for (var i = 1; i < lr.rows.length; i++) {
      var r = lr.rows[i] || [];
      var rowNo = lr.base + i + 1;
      var filled = r.some(function (v) { return cellText(v) !== ''; });
      if (!filled) continue;
      var lname = cellText(get(r, 'Investment line'));
      if (!lname) { messages.push({ level: 'error', row: rowNo, sheet: 'Lines', text: 'Row ' + rowNo + ': no investment line name. Row skipped.' }); continue; }
      var pre = 'Row ' + rowNo + ' (' + lname + '): ';
      function msg(level, text) { messages.push({ level: level, row: rowNo, sheet: 'Lines', text: pre + text }); }

      var rawAmt = get(r, 'Amount (SEK m)');
      var amount = fmt.parse(rawAmt instanceof Date ? 'x' : rawAmt);
      if (amount === null) amount = 0;
      if (isNaN(amount)) { msg('error', 'the amount "' + cellText(rawAmt) + '" is not a number. Row skipped.'); continue; }
      // SPEC5b D: Fixed or variable (blank = Fixed) and the rate of a variable line (its amount is then calculated on import)
      var rawBasis = cellText(get(r, 'Fixed or variable')), basis = 'Fixed';
      if (rawBasis) {
        if (norm(rawBasis) === 'variable') basis = 'Variable';
        else if (norm(rawBasis) !== 'fixed') msg('warning', '"Fixed or variable" is "' + rawBasis + '", not Fixed or Variable. Line imported as Fixed.');
      }
      var rate = null;
      if (basis === 'Variable') {
        var rawRate = get(r, 'Rate (%)');
        rate = fmt.parse(rawRate instanceof Date ? 'x' : typeof rawRate === 'string' ? rawRate.replace(/%\s*$/, '') : rawRate);
        if (rate !== null && isNaN(rate)) { msg('warning', 'the rate "' + cellText(rawRate) + '" is not a number. The rate is taken from the amount.'); rate = null; }
      }
      var rawType = get(r, 'Type');
      var type = parseType(rawType);
      if (!type) { msg('error', blank(cellText(rawType)) ? 'no type. Row skipped.' : 'the type "' + cellText(rawType) + '" is not Unconditional or Conditional. Row skipped.'); continue; }

      var pri = null, rawPri = cellText(get(r, 'Priority'));
      if (rawPri) {
        pri = findPriority(rawPri);
        if (pri === undefined) { pri = null; msg('warning', 'the priority "' + rawPri + '" is not in the priority list. Line imported with no priority.'); }
      }
      var rawTarget = get(r, 'Target');
      var target = fmt.parse(rawTarget instanceof Date ? 'x' : rawTarget);
      if (target !== null && isNaN(target)) { msg('warning', 'the target "' + cellText(rawTarget) + '" is not a number. Line imported without a target.'); target = null; }
      var dates = {};
      ['From', 'To'].forEach(function (h) {
        var raw = get(r, h), d = readDate(XLSX, raw);
        if (!d.ok) { msg('warning', 'the date "' + cellText(raw) + '" could not be read. Line imported without that date.'); dates[h] = null; }
        else dates[h] = d.value;
      });
      var rawActual = get(r, 'Actual to date');
      var actual = fmt.parse(rawActual instanceof Date ? 'x' : rawActual);
      if (actual !== null && isNaN(actual)) { msg('warning', 'the actual "' + cellText(rawActual) + '" is not a number. Line imported without an actual.'); actual = null; }
      var rawHow = cellText(get(r, 'Paid how')), paidHow = '';
      if (rawHow) {
        PAID_HOW.forEach(function (h) { if (!paidHow && norm(rawHow).indexOf(norm(h)) === 0) paidHow = h; });
        if (!paidHow) msg('warning', '"Paid how" is "' + rawHow + '", not On invoice, Off invoice or Fixed amount. Line imported without it.');
      }
      var rawBucket = cellText(get(r, 'Bucket')), bucket = '';
      if (rawBucket) {
        BUCKETS.forEach(function (b) { if (norm(rawBucket) === norm(b)) bucket = b; });
        if (!bucket) msg('warning', 'the bucket "' + rawBucket + '" is not Listings, Shelf space, In-store marketing, Promotions or Efficiency. Line imported without it.');
      }
      var rawPaid = get(r, 'Paid to date (SEK m)');
      var paid = fmt.parse(rawPaid instanceof Date ? 'x' : rawPaid);
      if (paid !== null && isNaN(paid)) { msg('warning', 'the paid to date "' + cellText(rawPaid) + '" is not a number. Line imported without it.'); paid = null; }

      var rawPA = get(r, 'Proposed amount (SEK m)');
      var propAmount = fmt.parse(rawPA instanceof Date ? 'x' : rawPA);
      if (propAmount !== null && isNaN(propAmount)) { msg('warning', 'the proposed amount "' + cellText(rawPA) + '" is not a number. Today\'s amount used.'); propAmount = null; }
      var rawPT = cellText(get(r, 'Proposed type'));
      var propType = rawPT ? parseType(rawPT) : type;
      if (!propType) { msg('warning', 'the proposed type "' + rawPT + '" is not Unconditional or Conditional. Today\'s type used.'); propType = type; }
      var rawPP = cellText(get(r, 'Proposed priority'));
      var propPri = pri;
      if (rawPP) {
        propPri = findPriority(rawPP);
        if (propPri === undefined) { propPri = null; msg('warning', 'the proposed priority "' + rawPP + '" is not in the priority list. Line proposed with no priority.'); }
      }
      // SPEC5 1: the new fields take their defaults; the action is set once the amounts are known (after TTS.recalc, below)
      lines.push({
        id: uid('l'), name: lname, amount: amount, paidHow: paidHow, bucket: bucket, type: type, priorityId: pri,
        propAmount: propAmount, propType: propType, propPriorityId: propPri,
        counterpart: cellText(get(r, 'Counterpart')), target: target, unit: cellText(get(r, 'Unit')),
        start: dates.From, end: dates.To, source: cellText(get(r, 'Data source')), owner: cellText(get(r, 'Owner')),
        paid: paid, clause: '', justification: '', action: 'Keep', why: '', priorityDecided: !!pri, basis: basis, rate: rate,
        plan: null, done: null, proof: 'None', tracked: false, checked: false, due: false, paidFull: false,
        actual: actual, kpiFull: cellText(get(r, 'KPI in full')), kpis: []
      });
    }

    // More KPIs
    var kpisRead = 0;
    var kws = sheetByName(wb, 'More KPIs');
    if (kws) {
      var kr = sheetRows(XLSX, kws);
      var kh = (kr.rows[0] || []).map(norm);
      var kc = {};
      KPI_HEAD.forEach(function (h) { kc[h] = kh.indexOf(norm(h)); });
      var kget = function (r, h) { var i = kc[h]; return i >= 0 ? r[i] : null; };
      for (var j = 1; j < kr.rows.length; j++) {
        var row = kr.rows[j] || [];
        var no = kr.base + j + 1;
        if (!row.some(function (v) { return cellText(v) !== ''; })) continue;
        var ln = cellText(kget(row, 'Investment line'));
        var kname = cellText(kget(row, 'KPI'));
        var kpre = 'Row ' + no + ' (' + (kname || ln || 'KPI') + '): ';
        var kmsg = function (level, text) { messages.push({ level: level, row: no, sheet: 'More KPIs', text: kpre + text }); };
        var line = null;
        lines.forEach(function (l) { if (!line && norm(l.name) === norm(ln)) line = l; });
        if (!line) { messages.push({ level: 'error', row: no, sheet: 'More KPIs', text: 'More KPIs, row ' + no + ': the investment line "' + ln + '" is not in the "Lines" sheet. KPI skipped.' }); continue; }
        var kt = kget(row, 'Target'), ktarget = fmt.parse(kt instanceof Date ? 'x' : kt);
        if (ktarget !== null && isNaN(ktarget)) { kmsg('warning', 'the target "' + cellText(kt) + '" is not a number. KPI imported without a target.'); ktarget = null; }
        var kd = {};
        ['From', 'To'].forEach(function (h) {
          var raw = kget(row, h), d = readDate(XLSX, raw);
          if (!d.ok) { kmsg('warning', 'the date "' + cellText(raw) + '" could not be read. KPI imported without that date.'); kd[h] = null; }
          else kd[h] = d.value;
        });
        var ka = kget(row, 'Actual to date'), kactual = fmt.parse(ka instanceof Date ? 'x' : ka);
        if (kactual !== null && isNaN(kactual)) { kmsg('warning', 'the actual "' + cellText(ka) + '" is not a number. KPI imported without an actual.'); kactual = null; }
        line.kpis.push({ id: uid('k'), name: kname, target: ktarget, unit: cellText(kget(row, 'Unit')), start: kd.From, end: kd.To,
          source: cellText(kget(row, 'Data source')), owner: cellText(kget(row, 'Owner')),
          plan: null, done: null, proof: 'None', tracked: false, checked: false, due: false, paidFull: false, actual: kactual,
          kpiFull: cellText(kget(row, 'KPI in full')) });
        kpisRead += 1;
      }
    }

    var channel = /away/i.test(cellText(fields.channel)) ? 'Away from Home' : 'Home';
    function fval(key) { var v = fmt.parse(fields[key] instanceof Date ? 'x' : fields[key]); return finite(v) ? v : null; }
    function fnum(key) { var v = fval(key); return v === null ? 0 : v; }
    // sales at list price; an older file gives "Revenue this year" (net invoice sales) instead: add the on-invoice lines back
    var listSales = fval('sales at list price (sek m)');
    if (listSales === null) {
      var onInv = lines.reduce(function (s, l) { return s + (l.paidHow === 'On invoice' ? l.amount : 0); }, 0);
      listSales = fval('revenue this year (sek m)') === null ? 0 : fval('revenue this year (sek m)') + onInv;
    }
    var customer = {
      id: null, name: name, channel: channel, example: false,
      sourceNote: 'Imported from "' + (fileName || 'file.xlsx') + '" on ' + fmt.date(isoOf(new Date())),
      listSales: listSales, revenueLY: fnum('revenue last year (sek m)'), investmentLY: fnum('investment last year (sek m)'),
      promoSpend: fnum('promotional spend (sek m)'),   // SPEC5b D; pack groups are not imported
      mustWinId: null, scriptEdits: {}, lines: lines
    };
    // monthly rows from the target, the period and the actual (SPEC5 3.4); the variable lines' amounts (TTS.recalc, SPEC5b A)
    normCustomer(customer, TTS.state.year, TTS.state.demoDate);
    // a blank proposed amount keeps today's (calculated) amount; the action follows from how the proposal differs from today
    lines.forEach(function (l) {
      if (l.propAmount === null) l.propAmount = l.amount;
      var action = 'Keep';
      if (r1(l.propAmount) !== r1(l.amount)) action = l.propAmount > l.amount ? 'Invest more' : 'Reduce';
      else if (l.propType !== l.type && l.propType === 'conditional') action = 'Make conditional';
      else if (l.propPriorityId !== l.priorityId) action = 'Move to a priority';
      l.action = action;
    });
    return { ok: true, customer: customer, linesRead: lines.length, kpisRead: kpisRead, messages: messages };
  }
  TTS.excel = excel;

  // Adds or replaces the customer from a parse result; returns the stored customer.
  TTS.applyImport = function (res) {
    var c = res.customer, st = S();
    var existing = null;
    st.customers.forEach(function (x) { if (norm(x.name) === norm(c.name)) existing = x; });
    if (existing) {
      c.id = existing.id;
      st.customers[st.customers.indexOf(existing)] = c;
    } else {
      var base = c.name.toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '') || 'customer';
      var id = base, n = 2;
      while (TTS.customer(id) || id === 'all') { id = base + '-' + n; n += 1; }
      c.id = id;
      st.customers.push(c);
    }
    return c;
  };

  // ---------------------------------------------------------------- routing
  function parseHash() {
    var h = HAS_DOM ? decodeURIComponent((location.hash || '').replace(/^#\/?/, '')) : '';
    var parts = h.split('/');
    return { screen: parts[0] || '', scope: parts[1] || '', option: parts[2] || '', sub: parts[3] || '', extra: parts.slice(4).join('/') };
  }
  function hashFor(screen, scope, option) { return '#' + screen + '/' + encodeURIComponent(scope) + (option ? '/' + option : ''); }
  TTS.hashFor = hashFor;
  function defaultScope() { var cs = S().customers; return cs.length ? cs[0].id : 'all'; }
  TTS.route = function () {
    var r = parseHash();
    var ok = SCREENS.some(function (s) { return s.id === r.screen; });
    var fixed = false;
    if (!ok) { r = { screen: 'priorities', scope: defaultScope(), option: '', sub: '', extra: '' }; fixed = true; }
    if (r.scope !== 'all' && !TTS.customer(r.scope)) { r.scope = defaultScope(); fixed = true; }
    r.packs = false;
    // step 3 (SPEC5 4, SPEC5c): one customer has two views, Contract analysis ('' or 'analysis', sub = a line id with its Details
    // open) and Agreement today ('waterfall', sub 'proposed'; 'waterfall/today/packs' and 'waterfall/proposed/packs' = by pack, only
    // for a customer with pack groups); All customers is the one chart ('' only). The old 'proposed' of one customer becomes
    // waterfall/proposed.
    if (r.screen === 'agreement') {
      if (r.scope === 'all') {
        if (r.option || r.sub) { r.option = ''; r.sub = ''; fixed = true; }
      } else if (r.option === 'proposed') {
        r.option = 'waterfall'; r.sub = 'proposed'; fixed = true;
      } else if (r.option === 'waterfall') {
        var cu = TTS.customer(r.scope);
        if ((r.sub === 'today' || r.sub === 'proposed') && r.extra === 'packs' && cu && Array.isArray(cu.packs) && cu.packs.length) { r.packs = true; r.extra = ''; }
        if (r.sub === 'today') { r.sub = ''; if (!r.packs) fixed = true; }
        if (r.sub && r.sub !== 'proposed') { r.sub = ''; fixed = true; }
      } else if (r.option !== 'analysis') {
        if (r.option !== '' || r.sub) { r.option = ''; r.sub = ''; fixed = true; }
      }
    }
    // Track & trace (SPEC5 7): #track/<c>/<kpiId> and #track/all/<c>/<kpiId> open that KPI's panel; an id that is not a KPI
    // of that customer (any more) is dropped; #track/<c> and #track/<c>/none = every panel closed (SPEC9 3)
    if (r.screen === 'track' && (r.option || r.sub)) {
      var tc = TTS.customer(r.scope === 'all' ? r.option : r.scope), tk = r.scope === 'all' ? r.sub : r.option;
      var okK = !!tc && !(r.scope !== 'all' && r.sub) && ((r.scope !== 'all' && tk === 'none') || E.kpis(tc).some(function (k) { return k.id === tk; }));
      if (!okK) { r.option = ''; r.sub = ''; fixed = true; }
    }
    if (r.extra) { r.extra = ''; fixed = true; }
    r.view = r.screen === 'agreement' && r.scope !== 'all' ? (r.option === 'waterfall' ? 'waterfall' : 'analysis') : '';
    r.fixed = fixed;
    return r;
  };
  TTS.go = function (screen, scope, option) {
    var h = hashFor(screen, scope, option);
    if (!HAS_DOM) return;
    if (location.hash === h) render(); else location.hash = h;
  };
  // the option to keep when moving to `screen` (toAll: the target scope is All customers). Agreement: from another screen
  // always Contract analysis (or All customers today); inside step 3 the Agreement today view and Today|Proposed are kept.
  function optionFor(screen, r, toAll) {
    // Track & trace keeps an open KPI panel across the scope switch: #track/a/a6 <-> #track/all/a/a6
    if (screen === 'track') {
      if (!r || r.screen !== 'track' || !r.option || r.option === 'none') return '';
      if (r.scope === 'all') return toAll ? r.option + '/' + r.sub : r.sub;
      return toAll ? r.scope + '/' + r.option : r.option;
    }
    if (screen !== 'agreement') return r && r.screen === screen ? r.option : '';
    if (!r || r.screen !== 'agreement') return '';
    var prop = r.option === 'proposed' || r.sub === 'proposed';
    if (toAll || r.scope === 'all') return '';
    return r.option === 'waterfall' ? 'waterfall' + (prop ? '/proposed' : '') : '';
  }
  function routeHash(r) {
    if (r.packs) return hashFor(r.screen, r.scope, r.option + '/' + (r.sub || 'today') + '/packs');
    return hashFor(r.screen, r.scope, r.option + (r.sub ? '/' + r.sub : ''));
  }

  // ---------------------------------------------------------------- shell rendering
  function topbar(r) {
    var st = S(), cs = st.customers;
    var anyExample = cs.some(function (c) { return c.example; });
    var isAll = r.scope === 'all';
    var shown = isAll ? (TTS.customer(st.ui.lastCustomerId) || cs[0]) : TTS.customer(r.scope);
    var opts = '';
    if (isAll && shown) opts += '<option value="" selected disabled hidden>' + esc(shown.name) + '</option>';
    cs.forEach(function (c) { opts += '<option value="' + esc(c.id) + '"' + (!isAll && c.id === r.scope ? ' selected' : '') + '>' + esc(c.name) + '</option>'; });
    var left = cs.length
      ? '<label class="scope-opt' + (isAll ? '' : ' on') + '"><select data-change="scope" data-key="scope" aria-label="Customer">' + opts + '</select>' + icon('caret-down', 16) + '</label>'
      : '<span class="scope-opt">No customers</span>';
    return '<header class="topbar"><div class="wrap">' +
      '<div class="brand"><span class="mark" aria-hidden="true"></span>Trade terms simulator</div>' +
      (anyExample ? '<span class="badge">Example data</span>' : '') +
      '<div class="scope" role="group" aria-label="Customer or all customers">' + left +
      '<a class="scope-opt' + (isAll ? ' on' : '') + '" href="' + hashFor(r.screen, 'all', optionFor(r.screen, r, true)) + '" data-key="scope-all">All customers</a></div>' +
      '<button type="button" class="btn onnavy" data-action="export" data-key="export">' + icon('download') + 'Export to Excel</button>' +
      '<button type="button" class="btn navytext" data-action="reset" data-key="reset" title="Reset to example data" aria-label="Reset to example data">' + icon('reset') + 'Reset</button>' +
      '</div></header>';
  }
  function chevrons(r) {
    return '<nav class="wrap" aria-label="Steps"><div class="steps">' + SCREENS.map(function (s, i) {
      var on = s.id === r.screen;
      return '<a class="step' + (on ? ' on' : '') + '" href="' + hashFor(s.id, r.scope, optionFor(s.id, r, r.scope === 'all')) + '"' + (on ? ' aria-current="page"' : '') +
        ' data-key="step-' + s.id + '"><span class="n">' + (i + 1) + '</span>' + esc(s.label) + '</a>';
    }).join('') + '</div></nav>';
  }
  // "Next: Import →" for the current route (keeps the scope); '' on the last step. Step 3 of one customer walks
  // Contract analysis → Agreement today → Track & trace (SPEC5 4).
  function nextButton(r) {
    r = r || TTS.route();
    function btn(href, label) { return '<a class="btn primary nextbtn" href="' + href + '" data-key="next">Next: ' + esc(label) + icon('arrow-right') + '</a>'; }
    if (r.view === 'analysis') return btn(hashFor('agreement', r.scope, 'waterfall'), 'Agreement today');
    var i = -1;
    SCREENS.forEach(function (s, j) { if (s.id === r.screen) i = j; });
    var nx = SCREENS[i + 1];
    if (i < 0 || !nx) return '';
    return btn(hashFor(nx.id, r.scope, optionFor(nx.id, r, r.scope === 'all')), nx.label);
  }
  // the row under a screen's content: optional left content (a ghost button or a note), the Next button on the right
  function bottomRow(left) {
    return '<div class="bottom-row">' + (left || '<span></span>') + nextButton() + '</div>';
  }
  function pageHead(h) {
    return '<div class="head"><div class="head-text"><h1>' + h.title + '</h1>' + (h.subtitle ? '<p class="sub">' + h.subtitle + '</p>' : '') + '</div>' +
      (h.actions ? '<div class="head-actions">' + h.actions + '</div>' : '') + '</div>';
  }

  function render() {
    if (!HAS_DOM) return;
    var app = document.getElementById('app');
    if (!app) return;
    var r = TTS.route();
    if (r.fixed && history.replaceState) history.replaceState(null, '', routeHash(r));
    var st = S();
    st.ui.screen = r.screen; st.ui.scope = r.scope;
    if (r.scope !== 'all') st.ui.lastCustomerId = r.scope;
    if (r.screen === 'agreement') st.ui.agreementMode = r.option === 'proposed' || r.sub === 'proposed' ? 'proposed' : 'today';
    // #agreement/<c>/analysis/<lineId> opens that line's Details once, when the route is reached, and brings it into view
    var openDet = null;
    if (r.view === 'analysis' && r.sub) {
      if (TTS.routeOpened !== location.hash) { TTS.routeOpened = location.hash; TTS.details[r.scope + ':' + r.sub] = true; openDet = r.sub; }
    } else TTS.routeOpened = null;

    // remember focus
    var ae = document.activeElement, fKey = TTS.focusKey || (ae && ae.getAttribute && ae.getAttribute('data-key')), sel = null;
    TTS.focusKey = null;
    if (ae && fKey && ae.getAttribute && ae.getAttribute('data-key') === fKey) {
      try { if (typeof ae.selectionStart === 'number') sel = [ae.selectionStart, ae.selectionEnd]; } catch (e) { sel = null; }
    }
    var sx = window.scrollX, sy = window.scrollY;

    var scr = TTS.screens[r.screen];
    var customer = r.scope === 'all' ? null : TTS.customer(r.scope);
    var head, body;
    if (!scr) {
      head = { title: esc(r.screen), subtitle: '' };
      body = emptyState({ icon: 'alert', title: 'This screen is missing', text: 'Its file did not load.' });
    } else {
      head = scr.head(r.scope, customer, r) || { title: '' };
      if (!st.customers.length && r.screen !== 'import') {
        head.actions = '';
        body = emptyState({ icon: 'file', title: 'No customers yet', text: 'Import a customer file or load the example data.',
          button: { label: 'Go to Import', href: hashFor('import', 'all') } });
      } else {
        body = customer ? scr.one(customer, r) : scr.all(r);
      }
    }
    // every screen but the last ends with the Next button; screens with their own bottom row put it there themselves
    if (body.indexOf('class="bottom-row') < 0 && nextButton(r)) body += bottomRow('');
    app.innerHTML = topbar(r) + chevrons(r) + '<main class="wrap page screen-' + r.screen + '">' + pageHead(head) +
      '<section class="view" id="view">' + body + '</section></main>';

    if (scr && scr.afterRender) { try { scr.afterRender(document.getElementById('view'), r); } catch (e) { if (global.console) console.error(e); } }
    renderPop();

    if (fKey) {
      var el = app.querySelector('[data-key="' + fKey.replace(/"/g, '\\"') + '"]') || document.querySelector('#tts-pop [data-key="' + fKey.replace(/"/g, '\\"') + '"]');
      if (el && el.focus) {
        el.focus({ preventScroll: true });
        if (sel && typeof el.setSelectionRange === 'function') { try { el.setSelectionRange(sel[0], sel[1]); } catch (e) { /* not a text input */ } }
      }
    }
    window.scrollTo(sx, sy);
    if (openDet) {
      var dr = app.querySelector('tr.ca-detrow[data-line="' + openDet.replace(/"/g, '\\"') + '"]');
      if (dr && dr.scrollIntoView) dr.scrollIntoView({ block: 'nearest' });
    }
    // a screen may ask for one element to be brought into view after this render (Track & trace: a panel just opened)
    if (TTS.scrollTarget) {
      var stEl = app.querySelector(TTS.scrollTarget);
      TTS.scrollTarget = null;
      if (stEl && stEl.scrollIntoView) stEl.scrollIntoView({ block: 'nearest' });
    }
  }
  TTS.render = render;
  TTS.commit = function () { TTS.recalc(); TTS.save(); render(); };

  // ---------------------------------------------------------------- period popover
  function popTarget() {
    var p = TTS.pop;
    if (!p) return null;
    var c = TTS.customer(p.cust), line = findLine(c, p.line);
    if (!line) return null;
    return p.kpi ? findKpi(line, p.kpi) : line;
  }
  function renderPop() {
    var layer = document.getElementById('tts-pop');
    if (!layer) return;
    var obj = popTarget();
    var anchor = obj && document.querySelector('[data-pop="' + TTS.pop.key + '"]');
    if (!obj || !anchor) { TTS.pop = null; layer.hidden = true; layer.innerHTML = ''; return; }
    var y = S().year;
    layer.innerHTML = '<div class="label">Period</div>' +
      '<div class="pop-fields"><label class="field"><span>From</span><input type="date" class="datefield" data-key="pop-from" value="' + esc(obj.start || '') + '" min="' + (y - 1) + '-01-01" max="' + (y + 1) + '-12-31"></label>' +
      '<label class="field"><span>To</span><input type="date" class="datefield" data-key="pop-to" value="' + esc(obj.end || '') + '" min="' + (y - 1) + '-01-01" max="' + (y + 1) + '-12-31"></label></div>' +
      '<div class="pop-actions"><button type="button" class="btn ghost" data-action="popWhole">' + icon('calendar') + 'Whole year</button>' +
      '<button type="button" class="btn ghost" data-action="popClear">Clear</button></div>';
    layer.hidden = false;
    var rc = anchor.getBoundingClientRect();
    var w = layer.offsetWidth;
    var left = Math.min(rc.left + window.scrollX, window.scrollX + document.documentElement.clientWidth - w - 16);
    layer.style.left = Math.max(16, left) + 'px';
    layer.style.top = (rc.bottom + window.scrollY + 6) + 'px';
  }
  function closePop(commitIt) {
    if (!TTS.pop) return;
    var obj = popTarget();
    if (commitIt && obj) {
      var f = document.querySelector('#tts-pop [data-key="pop-from"]'), t = document.querySelector('#tts-pop [data-key="pop-to"]');
      var was = [obj.start, obj.end];
      if (f) obj.start = f.value || null;
      if (t) obj.end = t.value || null;
      if (was[0] !== obj.start || was[1] !== obj.end) obj.plan = E.defaultPlan(obj);   // a new period gets the default plan
    }
    TTS.focusKey = TTS.pop.focusBack || null;
    TTS.pop = null;
    TTS.commit();
  }
  TTS.closePop = closePop;
  TTS.actions.period = function (el) {
    var key = el.getAttribute('data-pop');
    if (TTS.pop && TTS.pop.key === key) { closePop(true); return; }
    if (TTS.pop) closePop(true);
    TTS.pop = { key: key, cust: el.getAttribute('data-cust'), line: el.getAttribute('data-line'), kpi: el.getAttribute('data-kpi') || null,
      focusBack: el.getAttribute('data-key') };
    TTS.focusKey = 'pop-from';
    render();
  };
  TTS.actions.popWhole = function () {
    var y = S().year;
    var f = document.querySelector('#tts-pop [data-key="pop-from"]'), t = document.querySelector('#tts-pop [data-key="pop-to"]');
    if (f) f.value = y + '-01-01';
    if (t) t.value = y + '-12-31';
    closePop(true);
  };
  TTS.actions.popClear = function () {
    var f = document.querySelector('#tts-pop [data-key="pop-from"]'), t = document.querySelector('#tts-pop [data-key="pop-to"]');
    if (f) f.value = '';
    if (t) t.value = '';
    closePop(true);
  };

  // ---------------------------------------------------------------- shared actions and changes
  TTS.changes.scope = function (el) {
    var r = TTS.route();
    if (el.value) TTS.go(r.screen, el.value, optionFor(r.screen, r, el.value === 'all'));
  };
  // one field of a line or of an extra KPI (data-kpi). Besides storing the value: a new target, unit or period gives the KPI
  // its default plan (done stays); picking the priority marks it decided; picking an action applies its effect (SPEC5 3.2–3.4)
  TTS.changes.line = function (el) {
    var c = TTS.customer(el.getAttribute('data-cust'));
    var line = findLine(c, el.getAttribute('data-line'));
    if (!line) return;
    var kpiId = el.getAttribute('data-kpi');
    var obj = kpiId ? findKpi(line, kpiId) : line;
    if (!obj) return;
    var f = el.getAttribute('data-field'), v = el.value;
    if (el.getAttribute('data-num')) {
      var n = fmt.parse(f === 'rate' ? String(v).replace(/%\s*$/, '') : v);   // a rate may be typed with its "%"
      if (isNaN(n)) { toast('"' + v.trim() + '" is not a number.'); render(); return; }
      if (n === null && (f === 'amount' || f === 'propAmount' || f === 'rate')) n = 0;
      if (f === 'rate') n = r2(n);
      v = n;
    } else if (f === 'basis') {
      TTS.setBasis(c, line, v);   // Fixed | Variable; the amounts follow in TTS.commit (TTS.recalc)
      TTS.commit();
      return;
    } else if (f === 'priorityId' || f === 'propPriorityId') {
      v = E.priority(v) ? v : null;
      if (f === 'priorityId' && !kpiId) line.priorityDecided = true;
    } else if (f === 'action') {
      TTS.applyAction(line, v);
      TTS.commit();
      return;
    } else if (f === 'proof') {
      v = PROOFS.indexOf(v) >= 0 ? v : 'None';
    } else if (ANSWERS.indexOf(f) >= 0) {
      v = v === true || v === 'true' || v === 'yes';
    } else if (f === 'type' || f === 'propType') {
      v = v === 'conditional' ? 'conditional' : 'unconditional';
    } else if (f === 'start' || f === 'end') {
      v = v || null;
    } else if (f === 'paidHow') {
      v = PAID_HOW.indexOf(v) >= 0 ? v : '';
    } else if (f === 'bucket') {
      v = BUCKETS.indexOf(v) >= 0 ? v : '';
    } else {
      v = String(v).replace(/\s*[\r\n]+\s*/g, ' ').trim();   // wrapping editors never store a line break
      if (f === 'name' && !v && !kpiId) { toast('A contract point needs a name.'); render(); return; }
    }
    var replan = (f === 'target' || f === 'unit' || f === 'start' || f === 'end') && obj[f] !== v;
    // SPEC9 2.1: the KPI in full words must never contradict the numbers: a new counterpart, KPI name, target or unit clears it
    if ((f === (kpiId ? 'name' : 'counterpart') || f === 'target' || f === 'unit') && obj[f] !== v) obj.kpiFull = '';
    obj[f] = v;
    if (replan) obj.plan = E.defaultPlan(obj);
    TTS.commit();
  };
  // one customer figure: data-change="customer" data-cust data-field="listSales" | "promoSpend" | "revenueLY" | "investmentLY".
  // A number (comma accepted, bad text refused with a toast, blank = 0); promoSpend goes through TTS.setPromoSpend (rescales the
  // pack groups). The variable lines follow in TTS.commit.
  TTS.changes.customer = function (el) {
    var c = TTS.customer(el.getAttribute('data-cust')), f = el.getAttribute('data-field');
    if (!c || ['listSales', 'promoSpend', 'revenueLY', 'investmentLY'].indexOf(f) < 0) return;
    var n = fmt.parse(el.value);
    if (isNaN(n)) { toast('"' + String(el.value).trim() + '" is not a number.'); render(); return; }
    if (n === null) n = 0;
    if (f === 'promoSpend') TTS.setPromoSpend(c, n); else c[f] = n;
    TTS.commit();
  };
  // one figure of a pack group: data-change="pack" data-cust data-pack="<index in customer.packs>" data-field="share" | "promo".
  // A number (comma accepted, a "%" after a share accepted, bad text refused, blank = 0). The customer's promoSpend becomes the sum
  // of the pack groups' promo in TTS.commit. The shares are not forced to 100 (the view shows a note when they do not add up).
  TTS.changes.pack = function (el) {
    var c = TTS.customer(el.getAttribute('data-cust')), f = el.getAttribute('data-field');
    var p = c && Array.isArray(c.packs) ? c.packs[+el.getAttribute('data-pack')] : null;
    if (!p || (f !== 'share' && f !== 'promo')) return;
    var n = fmt.parse(String(el.value).replace(/%\s*$/, ''));
    if (isNaN(n)) { toast('"' + String(el.value).trim() + '" is not a number.'); render(); return; }
    p[f] = n === null ? 0 : Math.max(0, f === 'promo' ? r1(n) : n);
    TTS.commit();
  };
  // one month of a KPI: data-change="month" data-cust data-line [data-kpi] data-row="plan"|"done" data-month="0".."11".
  // Comma accepted; bad text refused with a toast; blank = 0 in the plan, nothing recorded in done.
  TTS.changes.month = function (el) {
    var line = findLine(TTS.customer(el.getAttribute('data-cust')), el.getAttribute('data-line'));
    var kpiId = el.getAttribute('data-kpi'), obj = kpiId ? findKpi(line, kpiId) : line;
    var row = el.getAttribute('data-row'), m = +el.getAttribute('data-month');
    if (!obj || (row !== 'plan' && row !== 'done') || !(m >= 0 && m < 12)) return;
    var n = fmt.parse(el.value);
    if (isNaN(n)) { toast('"' + String(el.value).trim() + '" is not a number.'); render(); return; }
    if (n === null && row === 'plan') n = 0;
    if (!Array.isArray(obj[row])) obj[row] = rep(row === 'plan' ? 0 : null);
    obj[row][m] = n;
    TTS.commit();
  };
  TTS.actions.export = function () {
    var r = TTS.route();
    // the formatted export (export.js, SPEC5 9) replaces this one when it is loaded
    if (typeof TTS.exportStyled === 'function') { TTS.exportStyled(r.scope); return; }
    if (typeof global.XLSX === 'undefined') { toast('The Excel library did not load.'); return; }
    try {
      var wb = excel.exportWorkbook(global.XLSX, S(), r.scope);
      global.XLSX.writeFile(wb, excel.exportName(S(), r.scope), { compression: true });
      toast('Excel file exported.');
    } catch (e) { toast('The Excel file could not be written.'); }
  };
  function doReset() {
    var r = TTS.route();
    TTS.resetToExample();
    var scope = r.scope === 'all' ? 'all' : 'a';
    TTS.go(r.screen, scope, optionFor(r.screen, r, scope === 'all'));
    render();
    toast('Example data loaded.');
  }
  TTS.actions.reset = function () {
    confirmDlg({ title: 'Reset to example data?', text: 'This replaces everything in the tool with the three example customers.', okLabel: 'Reset' })
      .then(function (ok) { if (ok) doReset(); });
  };
  TTS.actions.loadExample = function () {
    if (isUnchanged()) { doReset(); return; }
    TTS.actions.reset();
  };
  TTS.actions.go = function (el) {
    var r = TTS.route();
    TTS.go(el.getAttribute('data-screen') || r.screen, el.getAttribute('data-scope') || r.scope, el.getAttribute('data-option') || '');
  };

  // ================================================================ SCREEN 1 — Priorities
  function scopeName(scope, c) { return c ? c.name : 'all customers'; }
  TTS.screens.priorities = {
    head: function (scope, c) {
      return {
        title: 'The <span class="hl">must-win battles</span> for ' + esc(S().year),
        subtitle: c ? 'CCEP sets four to six priorities. Every investment at ' + esc(c.name) + ' is measured against them.'
          : 'CCEP sets four to six priorities. Every investment is measured against them.'
      };
    },
    one: function (c, r) { return prioritiesView(c.lines, c); },
    all: function () { return prioritiesView(TTS.scopeLines('all'), null); }
  };
  function spendCell(e, none) {
    return '<div class="spend' + (none ? ' none' : '') + '"><div class="spend-top"><span>' + fmt.money(e.amount) + '</span><span>' + fmt.pct(e.share) +
      '</span></div><div class="track"><i style="width:' + (Math.max(0, Math.min(1, e.share)) * 100).toFixed(2) + '%"></i></div></div>';
  }
  function prioritiesView(lines, c) {
    var st = S(), ps = st.priorities;
    var spend = E.spendByPriority(lines, 'today');
    var canDelete = ps.length > 4;
    var rows = ps.map(function (p, i) {
      var del = canDelete
        ? '<button type="button" class="iconbtn" data-action="deletePriority" data-id="' + esc(p.id) + '" data-key="prio:' + esc(p.id) + ':del" aria-label="Delete this priority" title="Delete this priority">' + icon('trash') + '</button>'
        : '<button type="button" class="iconbtn" aria-disabled="true" data-key="prio:' + esc(p.id) + ':del" aria-label="Delete this priority (at least four are needed)" title="At least four priorities are needed">' + icon('trash') + '</button>';
      return '<div class="prow">' +
        '<span class="pnum">' + (i + 1) + '</span>' +
        '<div class="ptext"><input class="cell pname" data-change="priority" data-id="' + esc(p.id) + '" data-field="name" data-key="prio:' + esc(p.id) + ':name" value="' + esc(p.name) + '" aria-label="Priority ' + (i + 1) + ' name" placeholder="Name of the priority">' +
        '<textarea rows="1" class="cell wrapedit pwin" data-change="priority" data-id="' + esc(p.id) + '" data-field="winning" data-key="prio:' + esc(p.id) + ':win" aria-label="What winning looks like" placeholder="What winning looks like">' + esc(p.winning) + '</textarea></div>' +
        spendCell(spend[i]) + '<div class="pdel">' + del + '</div></div>';
    }).join('');
    var none = spend[spend.length - 1];
    var noneRow = '<div class="prow none"><span class="pnum empty-num" aria-hidden="true"></span><div class="ptext"><div class="pname-static">No priority</div>' +
      '<div class="pwin-static">Spend that serves none of the priorities</div></div>' + spendCell(none, true) + '<div class="pdel"></div></div>';
    var add = ps.length < 6 ? '<button type="button" class="btn ghost addbtn" data-action="addPriority" data-key="prio-add">' + icon('plus') + 'Add a priority</button>' : '';
    var steps = ['Each investment line is linked to one priority, or to none.',
      'The negotiation plan shows the spend against each priority, today and proposed.',
      'Each customer gets one must-win battle, picked from this list.'];
    return '<div class="prio-layout"><div class="plist">' +
      '<div class="prow phead"><span></span><span class="label muted first">Priority and what winning looks like</span><span class="label muted">Spend today · ' + esc(scopeName(null, c)) + '</span><span></span></div>' +
      rows + noneRow + '</div>' +
      '<aside class="card explain"><div class="label">How the tool uses them</div><ol class="numlist">' +
      steps.map(function (s, i) { return '<li><span class="n">' + (i + 1) + '</span><span>' + esc(s) + '</span></li>'; }).join('') +
      '</ol></aside></div>' + bottomRow(add);
  }
  TTS.changes.priority = function (el) {
    var p = E.priority(el.getAttribute('data-id'));
    if (!p) return;
    var f = el.getAttribute('data-field'), v = String(el.value).replace(/\s*[\r\n]+\s*/g, ' ').trim();   // never store a line break
    if (f === 'name' && !v) { toast('A priority needs a name.'); render(); return; }
    p[f] = v;
    TTS.commit();
  };
  TTS.actions.addPriority = function () {
    var ps = S().priorities;
    if (ps.length >= 6) return;
    var p = { id: uid('p'), name: 'New priority', winning: '' };
    ps.push(p);
    TTS.focusKey = 'prio:' + p.id + ':name';
    TTS.commit();
  };
  TTS.actions.deletePriority = function (el) {
    var id = el.getAttribute('data-id'), st = S();
    if (st.priorities.length <= 4) return;
    confirmDlg({ title: 'Delete this priority?', text: 'Lines linked to it will have no priority.', okLabel: 'Delete' }).then(function (ok) {
      if (!ok) return;
      st.priorities = st.priorities.filter(function (p) { return p.id !== id; });
      st.customers.forEach(function (c) {
        if (c.mustWinId === id) c.mustWinId = null;
        c.lines.forEach(function (l) {
          if (l.priorityId === id) l.priorityId = null;
          if (l.propPriorityId === id) l.propPriorityId = null;
        });
      });
      TTS.commit();
      toast('Priority deleted.');
    });
  };

  // ================================================================ SCREEN 2 — Import
  TTS.screens['import'] = {
    head: function () {
      return { title: '<span class="hl">Import</span> the agreement of each customer', subtitle: 'One Excel file per customer, in the standard layout.' };
    },
    one: function (c) { return importView(c); },
    all: function () { return importView(null); },
    afterRender: function (root) {
      var zone = root.querySelector('.drop');
      if (!zone) return;
      zone.addEventListener('dragover', function (e) { e.preventDefault(); zone.classList.add('over'); });
      zone.addEventListener('dragenter', function (e) { e.preventDefault(); zone.classList.add('over'); });
      zone.addEventListener('dragleave', function (e) { if (!zone.contains(e.relatedTarget)) zone.classList.remove('over'); });
      zone.addEventListener('drop', function (e) {
        e.preventDefault(); zone.classList.remove('over');
        var f = e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0];
        if (f) importFile(f);
      });
    }
  };
  function importResult() {
    var res = TTS.lastImport;
    if (!res) return '';
    var h = '<div class="import-result">';
    if (res.ok) {
      var lr = res.linesRead, kr = res.kpisRead;
      var what = lr + (lr === 1 ? ' line' : ' lines') + (kr ? ' and ' + kr + (kr === 1 ? ' extra KPI' : ' extra KPIs') : '');
      h += '<div class="okband">' + icon('check', 20) + '<span>' + esc(res.customer.name) + ' imported from "' + esc(res.fileName) + '": ' + esc(what) + ' read.</span></div>';
    }
    if (res.messages.length) {
      h += '<ul class="msgs">' + (res.ok ? '' : '<li class="msg-lead">Nothing was imported.</li>') + res.messages.map(function (m) {
        var t = m.text.replace(/^More KPIs, row \d+: /, '').replace(/^Row \d+ \((.*?)\): /, '$1: ').replace(/^Row \d+: /, '');
        t = t.charAt(0).toUpperCase() + t.slice(1);
        var lab = m.row ? (m.sheet === 'More KPIs' ? 'More KPIs, row ' : 'Row ') + m.row : 'File';
        return '<li>' + chip('yellow', lab) + '<span>' + esc(t) + '</span></li>';
      }).join('') + '</ul>';
    }
    return h + '</div>';
  }
  function importView(c) {
    var st = S();
    var cards = '<div class="cards3">' +
      '<div class="card stepcard"><h2 class="cardhead"><span class="ncirc">1</span>Get the template</h2>' +
      '<p class="cardtext">Three sheets: Customer, Lines and More KPIs. One row per investment line.</p>' +
      '<div class="cardbtns"><button type="button" class="btn secondary" data-action="downloadTemplate" data-key="dl-template">' + icon('download') + 'Download the empty template</button>' +
      (c ? '<button type="button" class="btn ghost" data-action="downloadCustomer" data-id="' + esc(c.id) + '" data-key="dl-customer">' + icon('file') + 'Download ' + esc(c.name) + ' in this layout</button>' : '') +
      '</div></div>' +
      '<div class="card stepcard"><h2 class="cardhead"><span class="ncirc">2</span>Upload the customer\'s file</h2>' +
      '<label class="drop" data-key="drop">' + '<span class="drop-ic">' + icon('upload', 28) + '</span>' +
      '<span>Drop the Excel file here, or <span class="pick">choose a file</span></span>' +
      '<input type="file" class="visually-hidden" accept=".xlsx,.xls" data-change="importFile" data-key="file"></label></div>' +
      '<div class="card stepcard"><h2 class="cardhead"><span class="ncirc">3</span>Or start from the example</h2>' +
      '<p class="cardtext">Three invented customers: two Home, one Away from Home.</p>' +
      '<div class="cardbtns"><button type="button" class="btn primary" data-action="loadExample" data-key="load-example">Load example data</button></div></div>' +
      '</div>';
    var list = c ? [c] : st.customers;
    var rows = list.map(function (x) {
      var s = E.summary(x.lines, 'today');
      return '<tr><td class="strong">' + esc(x.name) + '</td><td>' + esc(x.channel) + '</td><td class="num">' + x.lines.length + '</td>' +
        '<td class="num strong">' + fmt.m(s.total) + '</td><td class="muted">' + (x.example ? '' : icon('file', 16) + ' ') + esc(x.sourceNote) + '</td>' +
        '<td class="act"><button type="button" class="iconbtn" data-action="removeCustomer" data-id="' + esc(x.id) + '" data-key="rm:' + esc(x.id) + '" aria-label="Remove ' + esc(x.name) + '" title="Remove ' + esc(x.name) + '">' + icon('trash') + '</button></td></tr>';
    }).join('');
    var total = '';
    if (!c && list.length) {
      var n = 0, t = 0;
      list.forEach(function (x) { n += x.lines.length; t += E.summary(x.lines, 'today').total; });
      total = '<tr class="total"><td>Total</td><td>' + list.length + (list.length === 1 ? ' customer' : ' customers') + '</td><td class="num">' + n + '</td><td class="num">' + fmt.m(t) + '</td><td></td><td></td></tr>';
    }
    var table = list.length
      ? '<table class="grid loaded"><colgroup><col style="width:22%"><col style="width:16%"><col style="width:9%"><col style="width:11%"><col><col style="width:64px"></colgroup>' +
        '<thead><tr><th>Customer</th><th>Channel</th><th class="num">Lines</th><th class="num">SEK <span class="nocase">m</span></th><th>Source</th><th><span class="visually-hidden">Remove</span></th></tr></thead>' +
        '<tbody>' + rows + total + '</tbody></table>'
      : '<p class="muted none-loaded">No customers are loaded. Upload a customer file or load the example data.</p>';
    return cards + importResult() + '<div class="loaded-wrap"><div class="label">What is loaded now</div>' + table + '</div>';
  }
  function downloadWb(build, name) {
    if (typeof global.XLSX === 'undefined') { toast('The Excel library did not load.'); return; }
    try { global.XLSX.writeFile(build(global.XLSX), name, { compression: true }); toast('Excel file downloaded.'); }
    catch (e) { toast('The Excel file could not be written.'); }
  }
  TTS.actions.downloadTemplate = function () {
    downloadWb(function (X) { return excel.templateWorkbook(X, S()); }, 'Trade terms - template.xlsx');
  };
  TTS.actions.downloadCustomer = function (el) {
    var c = TTS.customer(el.getAttribute('data-id'));
    if (!c) return;
    downloadWb(function (X) { return excel.customerWorkbook(X, S(), c); }, c.name.replace(/[\\/:*?"<>|]/g, '') + ' - standard layout.xlsx');
  };
  function importFile(file) {
    var name = file.name || 'file.xlsx';
    if (!/\.xlsx?$/i.test(name)) {
      TTS.lastImport = { ok: false, fileName: name, messages: [{ level: 'error', row: null, sheet: null, text: '"' + name + '" is not an Excel file. Use a .xlsx file in the standard layout.' }] };
      render(); return;
    }
    if (typeof global.XLSX === 'undefined') { toast('The Excel library did not load.'); return; }
    var reader = new FileReader();
    reader.onload = function () {
      var wb = null;
      try { wb = global.XLSX.read(new Uint8Array(reader.result), { type: 'array', cellDates: true }); } catch (e) { wb = null; }
      if (!wb) {
        TTS.lastImport = { ok: false, fileName: name, messages: [{ level: 'error', row: null, sheet: null, text: 'This file could not be read as an Excel file.' }] };
        render(); return;
      }
      var res = excel.parseCustomerWorkbook(global.XLSX, wb, S(), name);
      res.fileName = name;
      TTS.lastImport = res;
      if (res.ok) {
        var stored = TTS.applyImport(res);
        TTS.save();
        TTS.go('import', stored.id);
        toast(stored.name + ' imported.');
      } else {
        render();
      }
    };
    reader.onerror = function () { toast('The file could not be read.'); };
    reader.readAsArrayBuffer(file);
  }
  TTS.changes.importFile = function (el) {
    var f = el.files && el.files[0];
    if (f) importFile(f);
    el.value = '';
  };
  TTS.actions.removeCustomer = function (el) {
    var c = TTS.customer(el.getAttribute('data-id'));
    if (!c) return;
    confirmDlg({ title: 'Remove ' + c.name + '?', text: 'Its lines, KPIs and script are deleted.', okLabel: 'Remove' }).then(function (ok) {
      if (!ok) return;
      var st = S();
      st.customers = st.customers.filter(function (x) { return x !== c; });
      var r = TTS.route();
      TTS.save();
      if (r.scope === c.id) TTS.go(r.screen, st.customers.length ? 'all' : 'all', ''); else render();
      toast(c.name + ' removed.');
    });
  };

  // ================================================================ SCREEN 3 — Agreement (SPEC5 4 and 5)
  // One customer: two views behind a switch at the right of the page head, Contract analysis (default; route '' or 'analysis',
  // 'analysis/<lineId>' opens that line's Details) and Agreement today (route 'waterfall' or 'waterfall/proposed': the price
  // waterfall of Part 2, a stub for now). All customers: the table of the fourth pass until Part 2 replaces it.
  function segTab(label, href, on, key) {
    return '<a role="tab" class="' + (on ? 'on' : '') + '" aria-selected="' + on + '" href="' + href + '" data-key="' + key + '">' + label + '</a>';
  }
  // view 'all' = the chart: c is then the customer selected last (the first one if none)
  function viewSwitch(c, view) {
    return '<div class="segctl" role="tablist" aria-label="View">' + segTab('Contract analysis', hashFor('agreement', c.id, ''), view === 'analysis', 'view-analysis') +
      segTab('Agreement today', hashFor('agreement', c.id, 'waterfall'), view === 'waterfall', 'view-waterfall') +
      segTab('Compare customers', hashFor('agreement', 'all', ''), view === 'all', 'view-all') + '</div>';
  }
  TTS.screens.agreement = {
    head: function (scope, c, r) {
      if (!c) {
        var lc = TTS.customer(S().ui.lastCustomerId) || S().customers[0];
        return {
          title: 'Where the spend goes, <span class="hl">today and proposed</span>',
          subtitle: 'Each customer by bucket, and how the share linked to a priority moves.',
          actions: lc ? viewSwitch(lc, 'all') : ''
        };
      }
      if (r.view === 'waterfall') {
        var prop = r.sub === 'proposed', pk = r.packs ? '/packs' : '';
        return {
          title: 'The agreement <span class="hl">' + (prop ? 'proposed' : 'today') + '</span>, from gross revenue to net-net',
          subtitle: 'The same contract points as the Contract analysis, drawn as a price waterfall.',
          actions: viewSwitch(c, 'waterfall') + '<div class="segctl" role="tablist" aria-label="Today or proposed">' +
            segTab('Today', hashFor('agreement', c.id, 'waterfall' + (pk ? '/today' + pk : '')), !prop, 'mode-today') +
            segTab('Proposed', hashFor('agreement', c.id, 'waterfall/proposed' + pk), prop, 'mode-proposed') + '</div>'
        };
      }
      return {
        title: '<span class="hl">Contract analysis</span>: what the contract says, point by point',
        subtitle: 'Each contract point with its money, its counterpart and what we propose.',
        actions: viewSwitch(c, 'analysis')
      };
    },
    one: function (c, r) { return r.view === 'waterfall' ? waterfallView(c, r.sub === 'proposed' ? 'proposed' : 'today', r.packs) : analysisView(c); },
    all: function () { return agreementAll(); }
  };
  function attrs(c, line, kpi, field, extra) {
    return 'data-change="line" data-cust="' + esc(c.id) + '" data-line="' + esc(line.id) + '"' + (kpi ? ' data-kpi="' + esc(kpi.id) + '"' : '') +
      ' data-field="' + field + '" data-key="' + (kpi ? 'kpi:' + esc(c.id) + ':' + esc(kpi.id) : 'line:' + esc(c.id) + ':' + esc(line.id)) + ':' + field + '"' + (extra || '');
  }
  // wrapping editor: a one-row textarea that grows with its text (Enter commits, see TTS.start)
  function wrapEdit(c, line, kpi, field, value, ph, cls, label) {
    var cta = ph && ph.indexOf('Add ') === 0;
    return '<textarea rows="1" class="cell wrapedit' + (cls ? ' ' + cls : '') + (cta ? ' cta' : '') + '" ' + attrs(c, line, kpi, field) +
      ' placeholder="' + esc(ph || '') + '" aria-label="' + esc(label) + '">' + esc(value) + '</textarea>';
  }
  function textInput(c, line, kpi, field, value, ph, label, cls) {
    var cta = ph && ph.indexOf('Add ') === 0;
    return '<input class="cell' + (cls ? ' ' + cls : '') + (cta ? ' cta' : '') + '" ' + attrs(c, line, kpi, field) + ' value="' + esc(value) +
      '" placeholder="' + esc(ph || '') + '" aria-label="' + esc(label) + '">';
  }
  function dashCell() { return '<td class="dashcell"><span class="dash">–</span></td>'; }
  // number and unit side by side; "Add a target" sits on the number
  function targetInner(c, line, kpi, obj, cta) {
    var shown = finite(obj.target) ? fmt.num(obj.target) : '';
    return '<div class="tgt"><input class="cell num tnum' + (cta ? ' cta' : '') + '" ' + attrs(c, line, kpi, 'target', ' data-num="1"') +
      ' inputmode="decimal" value="' + esc(shown) + '" placeholder="' + (cta ? 'Add a target' : '') + '" aria-label="Target">' +
      wrapEdit(c, line, kpi, 'unit', obj.unit, kpi ? 'Unit' : '', 'tunit', 'Unit') + '</div>';
  }
  // a visible wrapping label with an invisible native select on top; the caret stays glued to the last word
  function ovLabel(c, line, field, label, text, muted, options, cls) {
    var words = String(text).split(' '), last = words.pop();
    var shown = (words.length ? esc(words.join(' ')) + ' ' : '') + '<span class="nw">' + esc(last) + '<span class="ovcaret">' + icon('caret-down', 12) + '</span></span>';
    return '<label class="cell ovsel' + (muted ? ' muted' : '') + (cls ? ' ' + cls : '') + '"><span class="ovtext">' + shown + '</span>' +
      '<select ' + attrs(c, line, null, field) + ' aria-label="' + esc(label) + '">' + options + '</select></label>';
  }
  function overlaySelect(c, line, field, label, text, muted, options, changed) {
    return '<td class="edit' + (changed ? ' changed' : '') + '">' + ovLabel(c, line, field, label, text, muted, options) + '</td>';
  }
  // bucket: shown and editable only when the line is conditional in this mode (like the counterpart), else a muted "–"
  function bucketCell(c, line, mode) {
    if (E.type(line, mode) !== 'conditional') return dashCell();
    var v = BUCKETS.indexOf(line.bucket) >= 0 ? line.bucket : '';
    var opts = '<option value=""' + (v ? '' : ' selected') + '>No bucket</option>' + BUCKETS.map(function (b) {
      return '<option value="' + b + '"' + (b === v ? ' selected' : '') + '>' + b + '</option>';
    }).join('');
    return overlaySelect(c, line, 'bucket', 'Bucket', v || 'No bucket', !v, opts);
  }
  function numInput(c, line, kpi, field, value, ph, label, money, cls) {
    var cta = ph && ph.indexOf('Add ') === 0;
    var shown = finite(value) ? (money ? fmt.m(value) : fmt.num(value)) : '';
    return '<input class="cell num' + (cls ? ' ' + cls : '') + (cta ? ' cta' : '') + '" ' + attrs(c, line, kpi, field, ' data-num="1"') +
      ' inputmode="decimal" value="' + esc(shown) + '" placeholder="' + esc(ph || '') + '" aria-label="' + esc(label) + '">';
  }
  function numCell(c, line, kpi, field, value, ph, label, money) {
    return '<td class="edit">' + numInput(c, line, kpi, field, value, ph, label, money) + '</td>';
  }
  // proposed amount with its change against today as a small second line (nothing when there is no change)
  function propAmountCell(c, line) {
    var d = r1(E.amount(line, 'proposed') - E.amount(line, 'today'));
    var change = d > 0 ? '<span class="chg up">' + fmt.delta(d) + '</span>' : d < 0 ? '<span class="chg down">' + fmt.delta(d) + '</span>' : '';
    return '<td class="edit"><div class="propamt">' + numInput(c, line, null, 'propAmount', line.propAmount, '0.0', 'Proposed amount in SEK m', true, 'bold') + change + '</div></td>';
  }
  // the classification chip with an invisible select on top that sets the type
  function typeCell(c, line, mode) {
    var field = mode === 'proposed' ? 'propType' : 'type';
    var t = E.type(line, mode);
    var lc = LINE_CHIP[E.classOf(line, mode)];
    return '<td class="typecell"><label class="typesel"><span class="chip ' + lc[0] + '">' + esc(lc[1]) + '<span class="tcaret">' + icon('caret-down', 12) + '</span></span>' +
      '<select ' + attrs(c, line, null, field) + ' aria-label="Classification">' +
      '<option value="unconditional"' + (t === 'unconditional' ? ' selected' : '') + '>Unconditional</option>' +
      '<option value="conditional"' + (t === 'conditional' ? ' selected' : '') + '>Conditional</option></select></label></td>';
  }
  function prioLabel(c, line, mode) {
    var field = mode === 'proposed' ? 'propPriorityId' : 'priorityId';
    var id = E.priorityId(line, mode);
    return ovLabel(c, line, field, mode === 'proposed' ? 'Proposed priority' : 'Priority', id ? E.priority(id).name : 'No priority', !id, priorityOptions(id));
  }
  function periodBtn(c, line, kpi, obj, cta) {
    var label = fmt.period(obj.start, obj.end);
    var key = (kpi ? 'kpi:' : 'line:') + c.id + ':' + (kpi ? kpi.id : line.id);
    var inner = label ? esc(label) : (cta ? '<span class="ctatext">Add a date</span>' : '<span class="dash">–</span>');
    var open = TTS.pop && TTS.pop.key === key;
    return '<button type="button" class="cell periodbtn' + (open ? ' open' : '') + '" data-action="period" data-pop="' + esc(key) + '" data-cust="' + esc(c.id) +
      '" data-line="' + esc(line.id) + '"' + (kpi ? ' data-kpi="' + esc(kpi.id) + '"' : '') + ' data-key="' + esc(key) + ':period" aria-label="When it runs' +
      (label ? ': ' + esc(label) : '') + '" aria-haspopup="dialog" aria-expanded="' + open + '">' + inner + '</button>';
  }
  // Details: which line rows are open (view state only, never saved; all closed on load)
  TTS.details = {};
  function detKey(c, line) { return c.id + ':' + line.id; }
  function detailsCell(c, line) {
    var open = !!TTS.details[detKey(c, line)];
    return '<td class="act"><button type="button" class="detbtn' + (open ? ' open' : '') + '" data-action="toggleDetails" data-cust="' + esc(c.id) +
      '" data-line="' + esc(line.id) + '" data-key="line:' + esc(c.id) + ':' + esc(line.id) + ':details" aria-expanded="' + open + '">Details' +
      icon('caret-down', 14) + '</button></td>';
  }
  var ORD = ['second', 'third', 'fourth', 'fifth', 'sixth'];
  function ordLabel(i) { var w = ORD[i] || 'extra'; return w.charAt(0).toUpperCase() + w.slice(1) + ' KPI'; }
  var SEKM = 'SEK <span class="nocase">m</span>';

  // ---------------------------------------------------------------- Contract analysis (SPEC5 5)
  function periodLower(start, end) {
    var p = fmt.period(start, end);
    return /^[A-Z][a-z]{2} to /.test(p) ? p : p.charAt(0).toLowerCase() + p.slice(1);
  }
  function missing(t) { return '<span class="ca-miss">' + esc(t) + '</span>'; }
  // the read-only summary of the internal counterpart under the name: counterpart · target · period (the owner is in Details;
  // only a missing owner shows here, as the yellow "no owner")
  function cpSummary(line) {
    if (E.type(line, 'today') !== 'conditional') return 'No counterpart in the contract';
    return [blank(line.counterpart) ? missing('no counterpart') : esc(line.counterpart),
      finite(line.target) ? esc(fmt.target(line.target, line.unit)) : missing('no number'),
      line.end ? esc(periodLower(line.start, line.end)) : missing('no date')].concat(blank(line.owner) ? [missing('no owner')] : [])
      .join('<span class="ca-sep"> · </span>');
  }
  var FLAG_SVG = '<svg class="ca-flag" width="16" height="16" viewBox="0 0 24 24" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true" focusable="false">' +
    '<path class="pole" d="M5.5 21V3.5"/><path d="M5.5 4h12l-2.6 4.25L17.5 12.5h-12z"/></svg>';
  // strength: four pips (Specific, Measurable, Owner, Time-bound) and, after a gap, the flag (aligned with a priority)
  function strengthIcon(line, mode) {
    if (E.type(line, mode) !== 'conditional') return '<span class="dash">–</span>';
    var s = E.smart(line, mode), al = E.aligned(line, mode);
    function yn(b) { return b ? 'yes' : 'no'; }
    var t = 'Specific: ' + yn(s.specific) + ' · Measurable: ' + yn(s.measurable) + ' · Owner: ' + yn(s.owner) + ' · Time-bound: ' + yn(s.timebound) +
      ' · Aligned with a priority: ' + yn(al);
    return '<span class="ca-str" role="img" title="' + esc(t) + '" aria-label="' + esc(t) + '"><span class="ca-pips">' +
      [s.specific, s.measurable, s.owner, s.timebound].map(function (x) { return '<i class="ca-pip' + (x ? ' on' : '') + '"></i>'; }).join('') +
      '</span><span class="ca-al' + (al ? ' on' : '') + '">' + FLAG_SVG + '</span></span>';
  }
  // priority today; while a suggestion is open: "Suggested: …" with a round confirm button (SPEC5 5)
  function caPrioCell(c, line) {
    var sug = E.suggestPriority(line), p = sug && E.priority(sug);
    if (!p) return '<td class="edit">' + prioLabel(c, line, 'today') + '</td>';
    var opts = '<option value="sug" selected disabled hidden>Suggested: ' + esc(p.name) + '</option>' + priorityOptions('-');
    // the confirm button floats at the right of the label (above the invisible select), so the text flows around it
    var ok = '<button type="button" class="ca-ok" data-action="confirmSuggestion" data-cust="' + esc(c.id) + '" data-line="' + esc(line.id) +
      '" data-key="line:' + esc(c.id) + ':' + esc(line.id) + ':sugok" aria-label="Confirm the suggested priority" title="Confirm the suggested priority: ' + esc(p.name) + '">' +
      icon('check', 16) + '</button>';
    return '<td class="edit">' + ovLabel(c, line, 'priorityId', 'Priority (suggested: ' + p.name + ')', 'Suggested: ' + p.name, false, opts, 'sug')
      .replace('<span class="ovtext">', '<span class="ovtext">' + ok) + '</td>';
  }
  function actionCell(c, line) {
    var opts = ACTIONS.map(function (a) { return '<option value="' + a + '"' + (a === line.action ? ' selected' : '') + '>' + a + '</option>'; }).join('');
    return '<td class="edit">' + ovLabel(c, line, 'action', 'Proposed action', line.action, line.action === 'Keep', opts) + '</td>';
  }
  // the % of gross revenue under an amount (SPEC5c 2): 13 px muted, one decimal
  function pgLine(pct) { return '<span class="ca-pg">' + fmt.pct1(pct) + '</span>'; }
  // SEK m · % of gross: a Fixed line types its amount, a Variable line shows the calculated amount as text (SPEC5b C); under it
  // the line's % of gross revenue (SPEC5c 2)
  function caAmountCell(c, line, row) {
    var top = line.basis !== 'Variable' ? numInput(c, line, null, 'amount', line.amount, '0.0', 'Amount in SEK m', true)
      : '<span class="ca-amtro" title="Calculated: the rate × its base">' + fmt.m(line.amount) + '</span>';
    return '<td class="edit ca-amtcell"><div class="ca-amt">' + top + pgLine(row.pct) + '</div></td>';
  }
  // Fixed or variable (SPEC5c 2): the overlay select on the first line; for a Variable line the rate input under it, with a title
  // naming its base
  function caBasisCell(c, line) {
    var variable = line.basis === 'Variable';
    var base = E.rateBase(line) === 'gross' ? 'of gross revenue' : 'of the first net';
    var opts = BASES.map(function (b) { return '<option value="' + b + '"' + (b === line.basis ? ' selected' : '') + '>' + b + '</option>'; }).join('');
    var rate = variable ? '<input class="cell ca-rin" ' + attrs(c, line, null, 'rate', ' data-num="1"') + ' inputmode="decimal" value="' + esc(fmt.rate(E.rateOf(line, c))) +
      '" placeholder="0.0%" title="' + base + '" aria-label="Rate in % ' + base + '">' : '';
    return '<td class="edit ca-fvcell"><div class="ca-fv">' + ovLabel(c, line, 'basis', 'Fixed or variable', variable ? 'Variable' : 'Fixed', false, opts, 'ca-basis') +
      rate + '</div></td>';
  }
  function caRow(c, line, row) {
    var flag = E.classOf(line, 'today') === 'partial';
    return '<tr class="ca-row ca-' + row.block + (flag ? ' flag' : '') + '" data-line="' + esc(line.id) + '">' +
      '<td class="edit ca-ptcell"><div class="ca-pt">' + wrapEdit(c, line, null, 'name', line.name, 'Name of the contract point', 'strong', 'Contract point') +
      '<div class="ca-sum">' + cpSummary(line) + '</div></div></td>' +
      caBasisCell(c, line) + caAmountCell(c, line, row) +
      typeCell(c, line, 'today') + bucketCell(c, line, 'today') + caPrioCell(c, line) +
      '<td class="ca-strcell">' + strengthIcon(line, 'today') + '</td>' +
      actionCell(c, line) + propAmountCell(c, line) + detailsCell(c, line) + '</tr>';
  }
  // the anchor rows (Gross revenue, the three nets) and the promotional spend row of the table (SPEC5c 2): label and hint over the
  // first two columns, the amount with its % of gross, the figure of the proposal; every other cell empty
  function caCustInput(c, field, value, label) {
    return '<input class="cell num ca-ain" data-change="customer" data-cust="' + esc(c.id) + '" data-field="' + field + '" data-key="cust:' + esc(c.id) + ':' + field +
      '" inputmode="decimal" value="' + esc(fmt.m(value)) + '" placeholder="0.0" aria-label="' + esc(label) + '">';
  }
  function caAnchorRow(c, row, prop) {
    var isPromo = row.kind === 'promo';
    var amt = row.key === 'gross' ? caCustInput(c, 'listSales', c.listSales, 'Gross revenue at list price in SEK m')
      : isPromo ? caCustInput(c, 'promoSpend', c.promoSpend, 'Promotional spend (estimate) in SEK m')
      : '<span class="ca-av">' + fmt.m(row.amount) + '</span>';
    var pv = prop[row.key], note = '';
    if (row.key === 'gross') {
      var folded = r1(row.amount - pv);
      if (folded > 0) note = '<span class="ca-fold">' + fmt.delta(-folded) + ' folded in</span>';
    }
    return '<tr class="' + (isPromo ? 'ca-promo' : 'ca-anchor ca-k-' + row.key) + '">' +
      '<td colspan="2" class="ca-alab"><div class="ca-an">' + esc(row.label) + '</div>' + (row.hint ? '<div class="ca-ah">' + esc(row.hint) + '</div>' : '') + '</td>' +
      '<td class="edit ca-amtcell"><div class="ca-amt">' + amt + pgLine(row.pct) + '</div></td><td colspan="5"></td>' +
      '<td class="ca-aprop"><div class="ca-apv">' + fmt.m(pv) + '</div>' + note + '</td><td></td></tr>';
  }
  // ---- the Details row: three blocks (In the contract · Internal counterpart · Proposal)
  function caField(label, control) {
    return '<div class="ca-f"><span class="ca-l">' + esc(label) + '</span><div class="ca-c">' + control + '</div></div>';
  }
  function paidHowSelect(c, line) {
    var v = PAID_HOW.indexOf(line.paidHow) >= 0 ? line.paidHow : '';
    var opts = '<option value=""' + (v ? '' : ' selected') + '>–</option>' + PAID_HOW.map(function (h) {
      return '<option value="' + h + '"' + (h === v ? ' selected' : '') + '>' + h + '</option>';
    }).join('');
    return '<select class="cell' + (v ? '' : ' muted') + '" ' + attrs(c, line, null, 'paidHow') + ' aria-label="Paid how">' + opts + '</select>';
  }
  function trashBtn(c, line, k) {
    return '<button type="button" class="iconbtn" data-action="deleteKpi" data-cust="' + esc(c.id) + '" data-line="' + esc(line.id) + '" data-kpi="' + esc(k.id) +
      '" data-key="kpi:' + esc(c.id) + ':' + esc(k.id) + ':del" aria-label="Delete this KPI" title="Delete this KPI">' + icon('trash') + '</button>';
  }
  function caKpi(c, line, k, i) {
    var m = E.missing(k, 'name');
    return '<div class="ca-kpi"><div class="ca-kpih"><span class="ca-kpin">' + esc(ordLabel(i)) + '</span>' + trashBtn(c, line, k) + '</div>' +
      // a small grid: name · target, unit, period · owner, data source
      '<div class="ca-kpig">' + wrapEdit(c, line, k, 'name', k.name, m.indexOf('counterpart') >= 0 ? 'Add a KPI name' : '', 'ca-kn', 'KPI name') +
      numInput(c, line, k, 'target', k.target, m.indexOf('target') >= 0 ? 'Add a target' : '', 'Target') + wrapEdit(c, line, k, 'unit', k.unit, 'Unit', '', 'Unit') +
      periodBtn(c, line, k, k, m.indexOf('date') >= 0) +
      textInput(c, line, k, 'owner', k.owner, 'Owner', 'Owner', 'ca-ko') + textInput(c, line, k, 'source', k.source, 'Data source', 'Data source', 'ca-ks') + '</div></div>';
  }
  function smartList(line, mode) {
    var s = E.smart(line, mode), pid = E.priorityId(line, mode);
    return '<div class="ca-smart"><div class="ca-cks">' + [['specific', 'Specific'], ['measurable', 'Measurable'], ['owner', 'With an owner'], ['timebound', 'Time-bound']].map(function (x) {
      return '<span class="ca-ck' + (s[x[0]] ? ' on' : '') + '">' + icon(s[x[0]] ? 'check' : 'cross', 16) + esc(x[1]) + '</span>';
    }).join('') + '</div><div class="ca-aln">Aligned with a priority: <b>' + (pid ? 'yes, ' + esc(E.priority(pid).name) : 'no') + '</b></div></div>';
  }
  // "Variable: 1.72% of the first net = SEK 3.6 m." / "Fixed: SEK 2.5 m = 1.2% of the first net." (SPEC5b C)
  function basisSentence(c, line) {
    var base = E.rateBase(line) === 'gross' ? 'of gross revenue' : 'of the first net', rate = E.rateOf(line, c);
    var r = finite(rate) ? fmt.rate(rate) : '–', m = fmt.money(E.amount(line, 'today')).replace(/ /g, ' ');   // "SEK 1.5 m" never breaks
    return line.basis === 'Variable' ? 'Variable: ' + r + ' ' + base + ' = ' + m + '.' : 'Fixed: ' + m + ' = ' + r + ' ' + base + '.';
  }
  function caDetails(c, line) {
    if (!TTS.details[detKey(c, line)]) return '';
    var ids = ' data-cust="' + esc(c.id) + '" data-line="' + esc(line.id) + '"', key = 'line:' + esc(c.id) + ':' + esc(line.id);
    var b1 = '<section class="ca-blk ca-b1"><h3 class="label">In the contract</h3>' +
      caField('Contract wording', wrapEdit(c, line, null, 'clause', line.clause, 'As written in the contract', '', 'Contract wording')) +
      caField('Justification found', wrapEdit(c, line, null, 'justification', line.justification, 'None stated', '', 'Justification found')) +
      caField('Paid how', paidHowSelect(c, line)) + '<p class="ca-basisnote">' + esc(basisSentence(c, line)) + '</p></section>';
    // the counterpart block follows today's type; a line unconditional today but conditional in the proposal edits it for the proposal
    var mode = E.type(line, 'today') === 'conditional' ? 'today' : E.type(line, 'proposed') === 'conditional' ? 'proposed' : null;
    var b2;
    if (!mode) {
      b2 = '<section class="ca-blk ca-b2"><h3 class="label">Internal counterpart</h3><p class="ca-nocp">No counterpart. Choose "Make conditional" to propose one.</p></section>';
    } else {
      var m = E.missing(line);
      var has = function (w) { return m.indexOf(w) >= 0; };
      b2 = '<section class="ca-blk ca-b2"><h3 class="label">Internal counterpart' + (mode === 'proposed' ? ' (for the proposal)' : '') + '</h3>' +
        caField('Counterpart', wrapEdit(c, line, null, 'counterpart', line.counterpart, has('counterpart') ? 'Add a counterpart' : '', '', 'Counterpart')) +
        caField('Target', targetInner(c, line, null, line, has('target'))) +
        caField('When it runs', periodBtn(c, line, null, line, has('date'))) +
        caField('Owner', textInput(c, line, null, 'owner', line.owner, has('owner') ? 'Add an owner' : '', 'Owner')) +
        caField('Data source', textInput(c, line, null, 'source', line.source, '–', 'Data source')) +
        ((line.kpis || []).length ? '<div class="ca-kpis">' + line.kpis.map(function (k, i) { return caKpi(c, line, k, i); }).join('') + '</div>' : '') +
        '<button type="button" class="btn ghost ca-addkpi" data-action="addKpi"' + ids + ' data-key="' + key + ':addkpi">' + icon('plus') + 'Add a KPI</button>' +
        smartList(line, mode) + '</section>';
    }
    var b3 = '<section class="ca-blk ca-b3"><h3 class="label">Proposal</h3>' +
      caField('Why', wrapEdit(c, line, null, 'why', line.why, 'One sentence on the proposal', '', 'Why')) +
      caField('Proposed priority', prioLabel(c, line, 'proposed')) +
      '<div class="ca-acts"><button type="button" class="btn ghost" data-action="deleteLine"' + ids + ' data-key="' + key + ':del">' + icon('trash') + 'Delete line</button></div></section>';
    return '<tr class="detrow ca-detrow" data-line="' + esc(line.id) + '"><td colspan="10"><div class="ca-det">' + b1 + b2 + b3 + '</div></td></tr>';
  }
  // the band: total, % of gross revenue, the classification (three small stats), SMART and aligned counterparts
  function caBand(c) {
    var kf = E.keyFigures(c), cl = kf.classes, cp = kf.counterparts;
    var mini = [['unconditional', 'u', 'Unconditional'], ['partial', 'n', 'Partially conditional'], ['conditional', 'c', 'Conditional']].map(function (x) {
      return '<div class="stat ca-mini' + (x[0] === 'conditional' ? ' good' : '') + '"><div class="v">' + Math.round(cl[x[0]].share * 100) + '<span class="u">%</span></div>' +
        '<div class="l"><i class="dot ' + x[1] + '"></i>' + x[2] + '</div></div>';
    }).join('');
    return '<div class="stats ca-band">' +
      '<div class="stat"><div class="v"><span class="u pre">SEK</span>' + fmt.m(kf.total) + '<span class="u">m</span></div><div class="l">Total trade terms</div></div>' +
      '<div class="stat"><div class="v">' + fmt.pct1(kf.pctOfGross).replace('%', '') + '<span class="u">%</span></div><div class="l">Of gross revenue</div></div>' +
      '<div class="ca-cls" role="group" aria-label="Classification">' + mini + '</div>' +
      '<div class="stat"><div class="v">' + cp.smart + '<span class="u">of ' + cp.lines + '</span></div><div class="l">Counterparts SMART</div></div>' +
      '<div class="stat"><div class="v">' + cp.aligned + '<span class="u">of ' + cp.lines + '</span></div><div class="l">Aligned with a priority</div></div></div>';
  }
  var CA_COLS = ['', 'ca-c-fv', 'ca-c-amt', 'ca-c-cls', 'ca-c-bkt', 'ca-c-pri', 'ca-c-str', 'ca-c-act', 'ca-c-prop', 'ca-c-det'];
  var STRENGTH_HELP = 'Four pips: Specific, Measurable, With an owner, Time-bound (filled green when met). Then a flag: aligned with a priority (filled green when it is).';
  var RATE_HELP = 'On-invoice lines: % of gross revenue. Other lines: % of the first net (net invoice sales). Variable: type the rate. Fixed: type the amount.';
  function qm(text, cls) { return '<span class="qm ' + cls + '" title="' + esc(text) + '" aria-label="' + esc(text) + '" role="img">?</span>'; }
  // the table reads top to bottom as the path from gross revenue to net-net: the rows of E.waterfall(c, 'today') (SPEC5c 2)
  function analysisView(c) {
    var head = ['Contract point', 'Fixed or' + qm(RATE_HELP, 'qm-rate') + '<br>variable', SEKM + '<br>% of gross', 'Classification', 'Bucket', 'Priority',
      'Strength<br class="br-narrow">' + qm(STRENGTH_HELP, 'qm-str'), 'Proposed<br>action', 'Proposed<br>' + SEKM];
    var nums = [2, 8];
    var colg = '<colgroup>' + CA_COLS.map(function (k) { return k ? '<col class="' + k + '">' : '<col>'; }).join('') + '</colgroup>';
    var thead = '<thead><tr>' + head.map(function (h, i) { return '<th' + (nums.indexOf(i) >= 0 ? ' class="num"' : '') + '>' + h + '</th>'; }).join('') +
      '<th><span class="visually-hidden">Details</span></th></tr></thead>';
    var prop = E.sales(c, 'proposed');
    var body = E.waterfall(c, 'today').map(function (row) {
      if (row.kind !== 'line') return caAnchorRow(c, row, prop);
      var l = findLine(c, row.lineId);
      return caRow(c, l, row) + caDetails(c, l);
    }).join('');
    var table = '<table class="grid agreement ca">' + colg + thead + '<tbody>' + body + '</tbody></table>';
    return caBand(c) + '<div class="table-gap ca-tbl">' + table + '</div>' +
      bottomRow('<button type="button" class="btn ghost addbtn" data-action="addLine" data-cust="' + esc(c.id) + '" data-key="line-add">' + icon('plus') + 'Add a contract point</button>');
  }

  // ---------------------------------------------------------------- Agreement today (SPEC5c 3): the picture of the same table
  // The band: total trade terms, % of gross revenue, the third net, then the two bucket bars (SPEC4) and their legend
  function wfBand(c, mode) {
    var s = E.sales(c, mode), t = mode === 'proposed' ? E.sales(c, 'today') : null;
    function stat(pre, v, unit, label, l2, today) {
      return '<div class="stat"><div class="v">' + (pre ? '<span class="u pre">' + pre + '</span>' : '') + v + '<span class="u">' + unit + '</span></div>' +
        '<div class="l">' + esc(label) + '</div>' + (l2 ? '<div class="l wf-l2">' + esc(l2) + '</div>' : '') +
        (today ? '<div class="l wf-today">' + esc(today) + '</div>' : '') + '</div>';
    }
    var pct = E.pctOfGross(s.tradeTerms, c, mode), tp = t ? E.pctOfGross(t.tradeTerms, c, 'today') : 0;
    return '<div class="stats bkband wf-band' + (t ? ' prop' : '') + '">' +
      stat('SEK', fmt.m(s.tradeTerms), 'm', 'Total trade terms', '', t ? 'today SEK ' + fmt.m(t.tradeTerms) + ' m' : '') +
      stat('', fmt.pct1(pct).replace('%', ''), '%', 'Of gross revenue', '', t ? 'today ' + fmt.pct1(tp) : '') +
      stat('SEK', fmt.m(s.thirdNet), 'm', 'Third net (net-net)', fmt.pct1(E.pctOfGross(s.thirdNet, c, mode)) + ' of gross revenue',
        t ? 'today SEK ' + fmt.m(t.thirdNet) + ' m' : '') +
      bucketChart(c.lines) + '</div>';
  }
  // the zoomed scale of the bars: from a little below the lowest point of the path (the third net) to gross revenue, the start
  // rounded down to a multiple of 5; { lo, hi } in SEK m
  function wfScale(rows) {
    var run = 0, lo = Infinity, hi = -Infinity;
    rows.forEach(function (r) {
      if (r.kind === 'anchor') run = r.amount; else run -= r.amount;
      lo = Math.min(lo, run); hi = Math.max(hi, run);
    });
    if (!(hi > lo)) return { lo: 0, hi: Math.max(hi, 1) };
    var start = Math.floor((lo - (hi - lo) * 0.1) / 5) * 5;
    if (lo >= 0 && start < 0) start = 0;
    return { lo: start, hi: hi };
  }
  var WF_ZERO = { 'Fold into the list price': 'folded into the list price', 'Remove': 'removed' };
  var BREAK = '<span class="wf-brk" aria-hidden="true"></span>';
  function waterfallBlock(c, mode, ctl) {
    var rows = E.waterfall(c, mode), sc = wfScale(rows), span = sc.hi - sc.lo;
    function x(v) { return Math.max(0, Math.min(100, (v - sc.lo) / span * 100)); }
    var guides = rows.filter(function (r) { return r.kind === 'anchor'; }).map(function (r) { return '<i class="wf-guide" style="left:' + x(r.amount).toFixed(3) + '%"></i>'; }).join('');
    function seg(from, to, cls, title) {
      var a = x(Math.min(from, to)), b = x(Math.max(from, to));
      return '<i class="wf-seg ' + cls + '" style="left:' + a.toFixed(3) + '%;width:' + Math.max(b - a, 0.25).toFixed(3) + '%" title="' + esc(title) + '"></i>';
    }
    var run = 0;
    var body = rows.map(function (r) {
      var before = run;
      run = r.kind === 'anchor' ? r.amount : run - r.amount;
      var amt = '<td class="num wf-amt">' + (r.kind === 'anchor' ? fmt.m(r.amount) : fmt.m(-r.amount)) + '</td><td class="num wf-pct">' + fmt.pct1(r.pct) + '</td>';
      if (r.kind === 'anchor') {
        var lab = r.hint;
        if (r.key === 'gross' && mode === 'proposed') {
          var folded = r1(E.sales(c, 'today').gross - r.amount);
          if (folded > 0) lab = 'SEK ' + fmt.m(folded) + ' m folded into the list price';
        }
        return '<tr class="wf-anchor wf-k-' + r.key + '"><td class="wf-lab"><div class="wf-an">' + esc(r.label) + '</div>' + (lab ? '<div class="wf-ah">' + esc(lab) + '</div>' : '') + '</td>' + amt +
          '<td class="wf-barcell"><div class="wf-track">' + guides + '<i class="wf-bar" style="width:' + x(r.amount).toFixed(3) + '%">' + BREAK + '</i></div></td></tr>';
      }
      if (r.kind === 'promo') {
        return '<tr class="wf-promo"><td class="wf-lab"><div class="wf-ln"><i class="dot wf-pdot"></i><span class="wf-n">' + esc(r.label) + '</span><span class="wf-tag">' + esc(r.hint) + '</span></div></td>' + amt +
          '<td class="wf-barcell"><div class="wf-track">' + guides + (r.amount > 1e-9 ? seg(before, run, 'wf-pseg', r.label + ': SEK ' + fmt.m(r.amount) + ' m, ' + fmt.pct1(r.pct) + ' of gross revenue') : '') + '</div></td></tr>';
      }
      var l = findLine(c, r.lineId), zero = Math.abs(r.amount) < 0.05;
      var tag = zero ? (mode === 'proposed' && WF_ZERO[l.action]) || 'no amount'
        : r.bucket + ' · ' + (r.basis === 'Variable' ? 'variable' + (mode === 'today' && finite(r.rate) ? ' ' + fmt.rate(r.rate) : '') : 'fixed');
      var href = hashFor('agreement', c.id, 'analysis/' + r.lineId);
      return '<tr class="wf-line click' + (zero ? ' wf-zero' : '') + '" data-action="go" data-screen="agreement" data-scope="' + esc(c.id) + '" data-option="analysis/' + esc(r.lineId) + '">' +
        '<td class="wf-lab"><div class="wf-ln"><i class="dot bk-' + esc(r.bucketKey) + '"></i><a class="wf-n" href="' + href + '" data-key="wf:' + esc(c.id) + ':' + esc(r.lineId) + '" title="Open in the Contract analysis">' +
        esc(r.label) + '</a><span class="wf-tag">' + esc(tag) + '</span>' + (r.cls === 'partial' && !zero ? '<span class="chip yellow wf-chip">Partially conditional</span>' : '') + '</div></td>' + amt +
        '<td class="wf-barcell"><div class="wf-track">' + guides + (zero ? '' : seg(before, run, 'bk-' + r.bucketKey, r.label + ': SEK ' + fmt.m(r.amount) + ' m, ' + fmt.pct1(r.pct) + ' of gross revenue')) + '</div></td></tr>';
    }).join('');
    var table = '<table class="wf" aria-label="Price waterfall, ' + (mode === 'proposed' ? 'proposed' : 'today') + '"><colgroup><col class="wf-c-lab"><col class="wf-c-amt"><col class="wf-c-pct"><col class="wf-c-bar"></colgroup>' +
      '<thead><tr><th class="wf-th0">' + (ctl || '<span class="visually-hidden">Row</span>') + '</th><th class="num">SEK <span class="nocase">m</span></th><th class="num">% of gross</th><th><span class="visually-hidden">Bar</span></th></tr></thead>' +
      '<tbody>' + body + '</tbody></table>';
    var cap = sc.lo > 0 ? 'The bars zoom in on the path from gross revenue to net-net: the scale starts at SEK ' + fmt.num(sc.lo) + ' m.' : '';
    return { html: table, caption: cap };
  }
  // By pack (SPEC5b G): one compact table from E.packWaterfall; each pack group's share and promotional spend are typed here
  var PACK_ROWS = [['gross', 'Gross revenue at list price', false, true], ['onInvoice', 'On invoice', true, false], ['firstNet', 'First net', false, true],
    ['offInvoice', 'Off invoice', true, false], ['secondNet', 'Second net', false, true], ['promo', 'Promotional spend (estimate)', false, false], ['thirdNet', 'Third net (net-net)', false, true]];
  function packBlock(c, mode, ctl) {
    var cols = E.packWaterfall(c, mode);
    var shareSum = c.packs.reduce(function (s, p) { return s + (finite(p.share) ? p.share : 0); }, 0);
    function packInput(i, field, value, label, text) {
      return '<input class="cell num wf-pin" data-change="pack" data-cust="' + esc(c.id) + '" data-pack="' + i + '" data-field="' + field + '" data-key="pack:' + esc(c.id) + ':' + i + ':' + field +
        '" inputmode="decimal" value="' + esc(text) + '" placeholder="0" aria-label="' + esc(label) + '">';
    }
    var head = '<tr><th class="wf-th0">' + (ctl || '<span class="visually-hidden">Row</span>') + '</th>' + cols.map(function (k) { return '<th class="num">' + esc(k.name) + '</th>'; }).join('') + '</tr>';
    var share = '<tr class="pk-share"><td class="pk-lab">Share of gross revenue</td>' + cols.map(function (k, i) {
      if (!i) return '<td class="num pk-v"><span class="pk-sum' + (Math.abs(shareSum - 100) > 0.05 ? ' off' : '') + '">' + fmt.num(r1(shareSum)) + '%</span></td>';
      var p = c.packs[i - 1];
      return '<td class="num edit">' + packInput(i - 1, 'share', p.share, 'Share of gross revenue, ' + p.name, fmt.num(p.share) + '%') + '</td>';
    }).join('') + '</tr>';
    var body = PACK_ROWS.map(function (pr) {
      var key = pr[0], neg = pr[2], net = pr[3];
      return '<tr class="' + (net ? 'pk-net' : 'pk-ded') + ' pk-' + key + '"><td class="pk-lab">' + esc(pr[1]) + '</td>' + cols.map(function (k, i) {
        var v = k[key], pc = k.pct[key];
        var top = key === 'promo' && i ? packInput(i - 1, 'promo', c.packs[i - 1].promo, 'Promotional spend, ' + k.name, fmt.m(c.packs[i - 1].promo))
          : '<span class="pk-amt">' + fmt.m(neg ? -v : v) + '</span>';
        var bar = key === 'thirdNet' ? '<span class="pk-bar"><i style="width:' + Math.max(0, Math.min(100, pc * 100)).toFixed(2) + '%"></i></span>' : '';
        return '<td class="num' + (key === 'promo' && i ? ' edit' : '') + '"><div class="pk-cell">' + top + '<span class="pk-pct">' + fmt.pct1(pc) + '</span>' + bar + '</div></td>';
      }).join('') + '</tr>';
    }).join('');
    var note = Math.abs(shareSum - 100) > 0.05 ? '<p class="pk-note">' + icon('alert', 16) + 'The pack groups\' shares add up to ' + fmt.num(r1(shareSum)) + '%, not 100%.</p>' : '';
    return '<table class="grid pk"><colgroup><col class="pk-c-lab">' + cols.map(function () { return '<col>'; }).join('') + '</colgroup><thead>' + head + '</thead><tbody>' + share + body + '</tbody></table>' + note;
  }
  function waterfallView(c, mode, packs) {
    var hasPacks = Array.isArray(c.packs) && c.packs.length > 0;
    var sub = mode === 'proposed' ? 'proposed' : 'today';
    var ctl = hasPacks ? '<div class="segctl small" role="tablist" aria-label="All packs or by pack">' +
      segTab('All packs', hashFor('agreement', c.id, 'waterfall' + (mode === 'proposed' ? '/proposed' : '')), !packs, 'packs-all') +
      segTab('By pack', hashFor('agreement', c.id, 'waterfall/' + sub + '/packs'), !!packs, 'packs-by') + '</div>' : '';
    if (packs && hasPacks) {
      return wfBand(c, mode) + '<div class="wf-block">' + packBlock(c, mode, ctl) + '</div>' +
        bottomRow('<p class="wf-cap">Contract points are spread across pack groups by revenue. Promotional spend is typed per pack group.</p>');
    }
    var wb = waterfallBlock(c, mode, ctl);
    return wfBand(c, mode) + '<div class="wf-block">' + wb.html + '</div>' + bottomRow(wb.caption ? '<p class="wf-cap">' + esc(wb.caption) + '</p>' : '');
  }

  // ---------------------------------------------------------------- All customers (SPEC5 6.2): one chart
  // For each customer and for all customers a pair of 100% columns, Today and Proposed, split by bucket (Unconditional at the
  // bottom); above each pair the share linked to a priority, today → proposed; each pair opens that customer's Contract analysis.
  var AC_DARK = { unconditional: 1, listings: 1, instore: 1, promotions: 1, nobucket: 1 };   // segments that take navy text
  function acColumn(lines, mode, label) {
    var es = E.spendByBucket(lines, mode).filter(function (e) { return e.amount > 1e-9; });
    var segs = es.map(function (e) {
      var tip = e.name + ': SEK ' + fmt.m(e.amount) + ' m, ' + fmt.pct(e.share);
      return '<i class="ac-seg bk-' + e.key + (AC_DARK[e.key] ? ' dk' : '') + '" style="flex-grow:' + e.share.toFixed(4) + '" title="' + esc(label + ' · ' + tip) + '">' +
        (e.share >= 0.08 ? '<b>' + fmt.pct(e.share) + '</b>' : '') + '</i>';
    }).join('');
    var aria = label + ': ' + (es.length ? es.map(function (e) { return e.name + ' ' + fmt.pct(e.share); }).join(', ') : 'no spend');
    return '<span class="ac-col' + (segs ? '' : ' empty') + '" role="img" aria-label="' + esc(aria) + '">' + segs + '</span>';
  }
  function acPair(name, channel, lines, href) {
    var lt = E.linkedShare(lines, 'today'), lp = E.linkedShare(lines, 'proposed');
    var a = Math.round(lt.share * 100), b = Math.round(lp.share * 100);
    var inner = '<span class="ac-hl"><span class="ac-hll">Linked to a priority</span><span class="ac-hlv">' + a + '%' + '<span class="ac-to">' + icon('arrow-right', 16) + '</span>' +
      '<span class="' + (b > a ? 'better' : '') + '">' + b + '%</span></span></span>' +
      '<span class="ac-cols">' + acColumn(lines, 'today', 'Today') + acColumn(lines, 'proposed', 'Proposed') + '</span>' +
      '<span class="ac-cl"><span>Today</span><span>Proposed</span></span>' +
      '<span class="ac-name">' + esc(name) + '</span><span class="ac-ch">' + esc(channel) + '</span>' +
      '<span class="ac-tot">SEK ' + fmt.m(lt.total) + ' m' + '<span class="ac-to">' + icon('arrow-right', 12) + '</span>' + fmt.m(lp.total) + ' m</span>';
    return href ? '<a class="ac-pair" href="' + href + '" data-key="pair:' + esc(href) + '" aria-label="' + esc(name) + ': open the contract analysis">' + inner + '</a>'
      : '<div class="ac-pair ac-all">' + inner + '</div>';
  }
  function agreementAll() {
    var st = S(), all = TTS.scopeLines('all');
    var chans = [];
    st.customers.forEach(function (c) { if (c.channel && chans.indexOf(c.channel) < 0) chans.push(c.channel); });
    var pairs = st.customers.map(function (c) { return acPair(c.name, c.channel, c.lines, hashFor('agreement', c.id, '')); }).join('') +
      '<span class="ac-sep" aria-hidden="true"></span>' + acPair('All customers', listJoin(chans), all, null);
    var t = E.spendByBucket(all, 'today'), p = E.spendByBucket(all, 'proposed');
    var legend = E.BUCKETS.filter(function (b, i) { return b.key !== 'nobucket' || t[i].amount > 1e-9 || p[i].amount > 1e-9; })
      .map(function (b) { return '<span><i class="dot bk-' + b.key + '"></i>' + esc(b.name) + '</span>'; }).join('');
    return '<div class="ac-chart"><div class="ac-pairs">' + pairs + '</div><div class="ac-legend">' + legend + '</div></div>';
  }

  // ---------------------------------------------------------------- step 3 actions
  TTS.actions.addLine = function (el) {
    var c = TTS.customer(el.getAttribute('data-cust'));
    if (!c) return;
    var l = normLine({ id: uid('l'), name: 'New contract point', amount: 0, paidHow: '', bucket: '', type: 'unconditional', priorityId: null,
      propAmount: 0, propType: 'unconditional', propPriorityId: null, counterpart: '', target: null, unit: '', start: null, end: null, source: '', owner: '',
      paid: null, clause: '', justification: '', action: 'Keep', why: '', priorityDecided: false, kpis: [] }, S().year, S().demoDate);
    c.lines.push(l);
    TTS.focusKey = 'line:' + c.id + ':' + l.id + ':name';
    TTS.commit();
  };
  TTS.actions.toggleDetails = function (el) {
    var cid = el.getAttribute('data-cust'), lid = el.getAttribute('data-line'), k = cid + ':' + lid;
    if (TTS.details[k]) {
      delete TTS.details[k];
      // closing the line the route opened: the route goes back to the plain Contract analysis
      var r = TTS.route();
      if (r.view === 'analysis' && r.sub === lid && history.replaceState) { history.replaceState(null, '', hashFor('agreement', cid, '')); TTS.routeOpened = null; }
    } else TTS.details[k] = true;
    TTS.focusKey = el.getAttribute('data-key');
    render();
  };
  TTS.actions.confirmSuggestion = function (el) {
    var c = TTS.customer(el.getAttribute('data-cust')), line = findLine(c, el.getAttribute('data-line'));
    var sug = line && E.suggestPriority(line);
    if (!sug) return;
    line.priorityId = sug;
    line.priorityDecided = true;
    TTS.focusKey = 'line:' + c.id + ':' + line.id + ':priorityId';
    TTS.commit();
  };
  TTS.actions.deleteLine = function (el) {
    var c = TTS.customer(el.getAttribute('data-cust')), line = findLine(c, el.getAttribute('data-line'));
    if (!line) return;
    confirmDlg({ title: 'Delete this contract point?', text: line.name + ' and its KPIs are removed from the agreement, today and proposed.', okLabel: 'Delete' }).then(function (ok) {
      if (!ok) return;
      c.lines = c.lines.filter(function (l) { return l !== line; });
      TTS.commit();
      toast('Contract point deleted.');
    });
  };
  TTS.actions.addKpi = function (el) {
    var c = TTS.customer(el.getAttribute('data-cust')), line = findLine(c, el.getAttribute('data-line'));
    if (!line) return;
    var k = normMonthly({ id: uid('k'), name: '', target: null, unit: '', start: null, end: null, source: line.source || '', owner: line.owner || '' }, S().year, S().demoDate);
    line.kpis = line.kpis || [];
    line.kpis.push(k);
    TTS.focusKey = 'kpi:' + c.id + ':' + k.id + ':name';
    TTS.commit();
  };
  TTS.actions.deleteKpi = function (el) {
    var c = TTS.customer(el.getAttribute('data-cust')), line = findLine(c, el.getAttribute('data-line'));
    var k = findKpi(line, el.getAttribute('data-kpi'));
    if (!k) return;
    confirmDlg({ title: 'Delete this KPI?', text: 'The line keeps its own KPI and its amount.', okLabel: 'Delete' }).then(function (ok) {
      if (!ok) return;
      line.kpis = line.kpis.filter(function (x) { return x !== k; });
      TTS.commit();
    });
  };

  // ---------------------------------------------------------------- start
  TTS.start = function () {
    TTS.load();
    if (!HAS_DOM) return;
    ['tts-pop', 'tts-toast'].forEach(function (id) {
      if (!document.getElementById(id)) {
        var d = document.createElement('div');
        d.id = id;
        if (id === 'tts-pop') { d.className = 'pop'; d.hidden = true; d.setAttribute('role', 'dialog'); d.setAttribute('aria-label', 'Period'); }
        else { d.className = 'toast'; d.setAttribute('role', 'status'); d.setAttribute('aria-live', 'polite'); }
        document.body.appendChild(d);
      }
    });
    if (!document.getElementById('tts-dialog')) {
      var dlg = document.createElement('dialog');
      dlg.id = 'tts-dialog'; dlg.className = 'dlg';
      document.body.appendChild(dlg);
    }
    document.addEventListener('click', function (e) {
      var t = e.target;
      if (TTS.pop && t.closest && !t.closest('#tts-pop') && !t.closest('[data-pop]')) {
        closePop(true);
        if (!(t.closest && t.closest('[data-action]'))) return;
      }
      var el = t.closest && t.closest('[data-action]');
      if (!el || el.getAttribute('aria-disabled') === 'true') return;
      var fn = TTS.actions[el.getAttribute('data-action')];
      if (!fn) return;
      if (el.tagName === 'A' && el.getAttribute('href') && el.getAttribute('href').charAt(0) === '#') e.preventDefault();
      if (el.tagName === 'TR' && t.closest('a, button, input, select, label')) return;
      fn(el, e);
    });
    document.addEventListener('change', function (e) {
      var el = e.target.closest && e.target.closest('[data-change]');
      if (!el) return;
      var fn = TTS.changes[el.getAttribute('data-change')];
      if (fn) fn(el, e);
    });
    document.addEventListener('keydown', function (e) {
      if (e.key === 'Escape' && TTS.pop) { e.preventDefault(); closePop(true); return; }
      if (e.key === 'Enter' && TTS.pop && e.target.closest && e.target.closest('#tts-pop input')) { e.preventDefault(); closePop(true); return; }
      // wrapping editors: Enter commits (focus is kept by the render) instead of adding a line break
      if (e.key === 'Enter' && !e.isComposing && e.target.classList && e.target.classList.contains('wrapedit')) {
        e.preventDefault();
        if (e.target.value !== e.target.defaultValue) e.target.dispatchEvent(new Event('change', { bubbles: true }));
        return;
      }
      if (e.key === 'Enter' && e.target.closest && e.target.closest('tr.click') && e.target.tagName === 'TR') {
        var fn = TTS.actions.go; fn(e.target, e);
      }
    });
    window.addEventListener('hashchange', function () { if (TTS.pop) TTS.pop = null; render(); });
    window.addEventListener('resize', function () { if (TTS.pop) renderPop(); });
    render();
  };

  global.TTS = TTS;
  return TTS;
})(typeof window !== 'undefined' ? window : globalThis);

// ====================================================================================================
/* Screen 4 — Track & trace. Registers into TTS.screens.track (SPEC 8.4, simplified in SPEC7). */
(function () {
  'use strict';
  var ui = TTS.ui, fmt = TTS.fmt, E = TTS.engine, esc = ui.esc;
  var MON = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  var MONTH = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
  var ORDER = { notexecuted: 0, partly: 1, executed: 2 };

  // columns (SPEC9 3): KPI (the rest of the width) · the year as twelve month cells · Status · chevron; widths in styles.css
  var DETH = '<span class="visually-hidden">Detail</span>';
  function cols() { return [['kpi', 'KPI'], ['cal', calHead()], ['st', 'Status'], ['ch', DETH]]; }

  function validIso(v) {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(String(v || ''))) return false;
    var p = v.split('-'), d = new Date(Date.UTC(+p[0], +p[1] - 1, +p[2]));
    return d.getUTCFullYear() === +p[0] && d.getUTCMonth() === +p[1] - 1 && d.getUTCDate() === +p[2];
  }
  function round1(n) { return Math.round(n * 10) / 10; }
  function periodLower(s) { return String(s || '').replace(/^(Whole|By|From)\b/, function (w) { return w.toLowerCase(); }); }
  function colgroup(cols) { return '<colgroup>' + cols.map(function (c) { return '<col class="tt-c-' + c[0] + '">'; }).join('') + '</colgroup>'; }
  function thead(cols) { return '<thead><tr>' + cols.map(function (c) { return '<th>' + c[1] + '</th>'; }).join('') + '</tr></thead>'; }
  function two(l1, l2, cls) { return '<span class="tt-l1' + (cls ? ' ' + cls : '') + '">' + l1 + '</span>' + (l2 ? '<span class="tt-l2">' + l2 + '</span>' : ''); }

  // ------------------------------------------------------------ cells
  // the state of one month (SPEC9 2.3), for the calendar in the table and the month cells of the panel:
  // done (reached) · part (recorded, below the plan) · late (planned, ended, nothing recorded) · plan (planned, still to come) · none
  function monthState(k, i, inP, ended) {
    var p = typeof k.plan[i] === 'number' ? k.plan[i] : 0, d = k.done[i];
    if (typeof d === 'number') {
      if (d >= p - 1e-9 && (p > 0 || d > 0)) return 'done';
      return p > 0 ? 'part' : 'none';
    }
    if (inP[i] && p > 0) return i < ended ? 'late' : 'plan';
    return 'none';
  }
  // the words of one month, as the title of its calendar cell
  function monthWords(k, i, st) {
    var pct = k.kind === 'percent', p = typeof k.plan[i] === 'number' ? k.plan[i] : 0, d = k.done[i];
    var pl = pct ? 'target ' + fmt.num(p) + '%' : fmt.num(p) + ' planned';
    var w = st === 'done' || st === 'part' ? pl + ', ' + (pct ? 'measured ' + fmt.num(d) + '%' : fmt.num(d) + ' executed')
      : st === 'late' ? pl + ', ' + (pct ? 'not measured' : 'not executed') : st === 'plan' ? pl : 'nothing planned';
    return MONTH[i] + ': ' + w;
  }
  // the calendar of one KPI (closed rows only): twelve coloured cells, no number inside; a light mark behind the demo month's
  // cell, only where that month has a cell (alone, the mark would read as a "still to come" month)
  function calCell(k) {
    var inP = E.periodMonths(k.start, k.end), ended = E.monthsEnded(), cur = E.currentMonth(), words = [], curHas = false;
    var cells = MON.map(function (m, i) {
      var st = monthState(k, i, inP, ended), w = monthWords(k, i, st);
      words.push(w);
      if (i === cur && st !== 'none') curHas = true;
      return '<i class="tt-d ' + st + '" title="' + esc(w) + '"></i>';
    }).join('');
    var mark = curHas ? '<span class="tt-nowmark" style="--cur:' + cur + '" aria-hidden="true"></span>' : '';
    return '<td class="tt-calcell">' + mark + '<span class="tt-cal" role="img" aria-label="' + esc(words.join('; ')) + '">' + cells + '</span></td>';
  }
  // the column header: the twelve month labels, each centred over its cell; the demo month is the navy pill
  function calHead() {
    var cur = E.currentMonth();
    return '<span class="tt-cal tt-calh">' + MON.map(function (m, i) {
      return '<span class="tt-ml' + (i === cur ? ' now' : '') + '"' + (i === cur ? ' title="Demo date ' + esc(fmt.date(TTS.state.demoDate)) + '"' : '') + '>' + m + '</span>';
    }).join('') + '</span>';
  }
  function legend() {
    return '<div class="tt-legend"><span><i class="tt-sw done"></i>Executed</span><span><i class="tt-sw late"></i>Planned, not executed</span>' +
      '<span><i class="tt-sw plan"></i>Still to come</span></div>';
  }
  // the KPI in words, then (muted) the contract point and the paid-ahead flag; All customers: the customer link first
  function aheadTitle(k) {
    return fmt.pct(k.paidShare) + ' of the money is paid, ' + fmt.pct(k.lineDelivered) + ' of the target is delivered.';
  }
  function kpiText(k, scopeAll) {
    var flag = k.ahead ? '<span class="tt-ahead" title="' + esc(aheadTitle(k)) + '">' + ui.icon('alert', 14) + '<span>Paid ahead of delivery</span></span>' : '';
    var where = (scopeAll ? '<a class="tt-cust" href="' + TTS.hashFor('track', k.customerId) + '" data-key="track-cust:' + esc(k.customerId) + ':' + esc(k.id) + '">' +
      esc(k.customerName) + '</a> · ' : '') + esc(k.lineName);
    return two(esc(k.text), '<span>' + where + '</span>' + flag, 'tt-name');
  }
  function kpiCell(k, scopeAll) { return '<td class="tt-kpi">' + kpiText(k, scopeAll) + '</td>'; }

  // ------------------------------------------------------------ stat band and footer
  function band(kpis, customers) {
    var n = kpis.length, cnt = { executed: 0, partly: 0, notexecuted: 0 };
    var ahead = E.paidAheadTotal(customers);
    kpis.forEach(function (k) { cnt[k.status] += 1; });
    var chips = [['executed', 'green', 'Fully executed'], ['partly', 'yellow', 'Partially executed'], ['notexecuted', 'outline', 'Not executed']].map(function (x) {
      return '<span class="chip ' + x[1] + ' tt-bigchip' + (cnt[x[0]] ? '' : ' zero') + '"><b>' + cnt[x[0]] + '</b>' + x[2] + '</span>';
    }).join('');
    var d = TTS.state.demoDate;
    var right = '<div class="tt-status"><div class="tt-chips">' + chips + '</div><div class="l">Status on ' + esc(validIso(d) ? fmt.date(d) : '') + '</div></div>';
    return ui.statBand([
      { value: String(n), label: n === 1 ? 'KPI to track' : 'KPIs to track' },
      { pre: 'SEK', value: fmt.m(E.moneyAtStake(kpis)), unit: 'm', label: n === 1 ? 'Depends on it' : 'Depends on them' },
      { pre: 'SEK', value: fmt.m(ahead), unit: 'm', label: 'Paid ahead of delivery', cls: fmt.m(ahead) === '0.0' ? 'tt-nonah' : 'tt-ahst' }
    ], right);
  }
  function noKpiNote(c) {
    var s = E.summary(c.lines, 'today');
    var parts = [];
    if (s.unconditional > 0.0001) parts.push(fmt.money(s.unconditional) + ' unconditional');
    if (s.partial > 0.0001) parts.push(fmt.money(s.partial) + ' partially conditional');
    var rest = s.unconditional + s.partial;
    if (!parts.length || Math.round(rest * 10) === 0) return '<div></div>';
    return '<p class="tt-note">' + ui.icon('alert') + '<span>' + esc(fmt.money(rest)) + ' has no KPI: ' + esc(parts.join(' and ')) + '.</span>' +
      '<a class="linkbtn tt-open" href="' + TTS.hashFor('agreement', c.id) + '" data-key="track-open-agreement">Open the agreement' + ui.icon('arrow-right', 16) + '</a></p>';
  }
  function emptyKpis(scope) {
    return ui.emptyState({ icon: 'calendar', title: 'No KPIs yet', text: 'A line becomes a KPI when it has a counterpart, a target number and a date.',
      button: { label: 'Open the agreement', href: TTS.hashFor('agreement', scope) } });
  }

  // ------------------------------------------------------------ drill-down: chevron, route, panel
  var PAID_HOW = ['On invoice', 'Off invoice', 'Fixed amount'];
  var SWITCHES = [['tracked', 'The KPI is tracked'], ['checked', 'Execution has been checked'], ['due', 'The payment is already due'], ['paidFull', 'It has been paid in full']];
  var editing = null;   // "<customerId>:<kpiId>" whose Month by month shows the number inputs (view state, never saved)
  // the KPI whose panel is open: #track/<c>/<kpiId> (TTS.route has already dropped an id that is not a KPI of that customer)
  function openFrom(r, cid) { return r && r.scope === cid ? r.option : ''; }
  // closing: All customers -> #track/all; one customer -> #track/<c> (every panel closed, SPEC9 3)
  function trackHash(k, scopeAll, open) {
    if (open) return scopeAll ? TTS.hashFor('track', 'all') : TTS.hashFor('track', k.customerId);
    return TTS.hashFor('track', scopeAll ? 'all' : k.customerId, scopeAll ? k.customerId + '/' + k.id : k.id);
  }
  // bring a panel into view once, when its route is reached (clicking a row or opening the URL)
  function scrollOnOpen(r) {
    if (r && r.option && r.option !== 'none' && typeof location !== 'undefined') {
      if (TTS.trackOpened !== location.hash) { TTS.trackOpened = location.hash; TTS.scrollTarget = 'tr.tt-detrow'; }
    } else TTS.trackOpened = null;
  }
  function keyOf(k, what) { return 'track:' + k.customerId + ':' + k.id + ':' + what; }
  function chevBtn(k, open, scopeAll) {
    return '<button type="button" class="iconbtn tt-chev' + (open ? ' open' : '') + '" data-action="trackDetails" data-href="' +
      esc(trackHash(k, scopeAll, open)) + '" data-key="' + esc(keyOf(k, 'details')) + '" aria-expanded="' + open + '" aria-label="' +
      (open ? 'Hide the detail' : 'Show the detail') + '">' + ui.icon('caret-down', 20) + '</button>';
  }
  // a closed KPI row: the whole row opens its card (a click on a link or button in the row keeps its own behaviour: see TTS.start)
  function closedRow(k, scopeAll) {
    return '<tr class="tt-krow" data-action="trackDetails" data-href="' + esc(trackHash(k, scopeAll, false)) + '">' + kpiCell(k, scopeAll) + calCell(k) +
      '<td>' + ui.statusChip(k.status) + '</td><td class="tt-chcell">' + chevBtn(k, false, scopeAll) + '</td></tr>';
  }
  TTS.actions.trackDetails = function (el, e) {
    // the card head closes the card, like the row opens it: a click on a link or button in it keeps its own behaviour
    if (el.tagName === 'DIV' && e && e.target.closest('a, button, input, select, label')) return;
    var h = el.getAttribute('data-href');
    if (el.tagName === 'BUTTON') TTS.focusKey = el.getAttribute('data-key');
    if (location.hash === h) TTS.render(); else location.hash = h;
  };
  // the stored object behind a KPI: the line itself (its own KPI) or the extra KPI in line.kpis
  function lineOf(c, k) { return c ? c.lines.filter(function (l) { return l.id === k.lineId; })[0] : null; }
  function objOf(c, k) {
    var line = lineOf(c, k);
    if (!line || !k.isExtra) return line;
    return (line.kpis || []).filter(function (x) { return x.id === k.id; })[0] || null;
  }
  // the attributes that point a control at this KPI (own KPI: the line; extra KPI: data-kpi)
  function ids(k) {
    return 'data-cust="' + esc(k.customerId) + '" data-line="' + esc(k.lineId) + '"' + (k.isExtra ? ' data-kpi="' + esc(k.id) + '"' : '');
  }
  var NBSP = ' ';

  // block 1 (SPEC7 B1, SPEC9 4): plain text and a link to where the KPI, its data source and its owner are edited
  function agreedBlock(k) {
    var per = periodLower(fmt.period(k.start, k.end));
    var what = k.text;
    var src = [];
    if (String(k.source || '').trim()) src.push('Tracked through: ' + k.source);
    if (String(k.owner || '').trim()) src.push('Owner: ' + k.owner);
    return '<section class="ca-blk ca-b1"><h3 class="label">What was agreed</h3>' +
      '<p class="tt-agreed">' + esc(what + (per ? ', ' + per : '') + '.') + '</p>' +
      '<p class="tt-money"><b>' + esc(fmt.money(k.money).replace(/ /g, NBSP)) + '</b> depends on it.</p>' +
      (src.length ? '<p class="tt-src">' + esc(src.join(' · ')) + '</p>' : '') +
      '<a class="linkbtn tt-edlink" href="' + TTS.hashFor('agreement', k.customerId, 'analysis/' + k.lineId) + '" data-key="' + esc(keyOf(k, 'analysis')) + '">Edit in the Contract analysis</a></section>';
  }

  // block 2 (SPEC7 B2): one row of twelve month cells; "Edit numbers" swaps it for two rows of inputs
  function monthCell(k, i, inP, ended, cur) {
    var live = i < ended || i === cur;
    var pct = k.kind === 'percent', p = typeof k.plan[i] === 'number' ? k.plan[i] : 0, d = k.done[i];
    var rec = typeof d === 'number', st = monthState(k, i, inP, ended), txt;
    if (st === 'none') return '<span class="tt-mc none" aria-hidden="true">–</span>';
    txt = !rec ? fmt.num(p) : st === 'part' && !pct ? fmt.num(d) + '/' + fmt.num(p) : fmt.num(d);
    var sm = txt.length >= (st === 'part' ? 4 : 5) ? ' sm' : '';
    var tick = st === 'done' ? '<span class="tt-tk">' + ui.icon('check', 12) + '</span>' : '';
    var cls = 'tt-mc ' + st + sm;
    if (!live) return '<span class="' + cls + '">' + esc(txt) + tick + '</span>';
    var lab = MONTH[i] + ': ' + (pct ? 'target ' + fmt.num(p) + '%, ' + (rec ? 'measured ' + fmt.num(d) + '%' : 'not measured')
      : fmt.num(p) + ' planned, ' + (!rec ? 'not executed' : st === 'done' ? 'executed' : fmt.num(d) + ' executed'));
    return '<button type="button" class="' + cls + '" data-action="trackMonth" ' + ids(k) + ' data-month="' + i + '" data-key="' + esc(keyOf(k, 'm:' + i)) +
      '" aria-pressed="' + rec + '" aria-label="' + esc(lab) + '">' + esc(txt) + tick + '</button>';
  }
  function monthInput(k, row, i) {
    var pct = k.kind === 'percent', v = k[row][i];
    return '<input class="tt-mi" type="text" inputmode="decimal" data-change="month" ' + ids(k) + ' data-row="' + row + '" data-month="' + i +
      '" data-key="' + esc(keyOf(k, row + ':' + i)) + '" value="' + esc(typeof v === 'number' ? fmt.num(v) : '') + '" aria-label="' +
      (row === 'plan' ? (pct ? 'Target' : 'Planned') : (pct ? 'Measured' : 'Executed')) + ', ' + MONTH[i] + '">';
  }
  function monthsBlock(k) {
    var pct = k.kind === 'percent', inP = E.periodMonths(k.start, k.end), ended = E.monthsEnded(), cur = E.currentMonth();
    var id = k.customerId + ':' + k.id, edit = editing === id;
    var head = '<div class="tt-mhead"><h3 class="label">Month by month</h3><button type="button" class="linkbtn tt-edit" data-action="trackEdit" data-id="' + esc(id) +
      '" data-key="' + esc(keyOf(k, 'edit')) + '">' + (edit ? 'Done' : 'Edit numbers') + '</button></div>';
    var labels = MON.map(function (m, i) {
      return '<span class="tt-ml' + (i === cur ? ' now' : '') + '"' + (i === cur ? ' title="Demo date ' + esc(fmt.date(TTS.state.demoDate)) + '"' : '') + '>' + m + '</span>';
    }).join('');
    var body;
    if (edit) {
      body = '<span class="tt-rl">' + (pct ? 'Target' : 'Planned') + '</span>' + MON.map(function (m, i) { return monthInput(k, 'plan', i); }).join('') +
        '<span class="tt-rl">' + (pct ? 'Measured' : 'Executed') + '</span>' + MON.map(function (m, i) { return monthInput(k, 'done', i); }).join('') +
        // the same two row labels again for July to December, shown only when the strip wraps into two rows of six
        '<span class="tt-rl tt-rl2" aria-hidden="true">' + (pct ? 'Target' : 'Planned') + '</span><span class="tt-rl tt-rl2" aria-hidden="true">' + (pct ? 'Measured' : 'Executed') + '</span>';
    } else {
      body = MON.map(function (m, i) { return monthCell(k, i, inP, ended, cur); }).join('');
    }
    var sum;
    if (pct) {
      sum = (k.actual === null ? 'Nothing measured yet' : 'Latest measured: ' + fmt.num(k.actual) + '%') + ' · target ' + fmt.num(k.target) + '%';
    } else {
      sum = 'Executed ' + fmt.num(k.actual || 0) + ' of ' + fmt.num(round1(k.expected || 0)) + ' planned so far · ' + fmt.num(k.target) + ' agreed' +
        (fmt.period(k.start, k.end) === 'Whole year' ? ' for the year' : '');
    }
    return '<section class="ca-blk ca-b2"><div class="tt-mm">' + head + '<div class="tt-strip' + (edit ? ' edit' : '') + '">' + labels + body + '</div>' +
      '<p class="tt-msum">' + esc(sum) + '</p></div></section>';
  }

  // block 3 (SPEC7 B3): two drop-downs and "Paid so far" on one label grid, then the four answers as switches
  function selectOf(k, field, value, label, options, blank) {
    var v = options.indexOf(value) >= 0 ? value : '';
    var opts = (blank ? '<option value=""' + (v ? '' : ' selected') + '>–</option>' : '') + options.map(function (o) {
      return '<option value="' + esc(o) + '"' + (o === v ? ' selected' : '') + '>' + esc(o) + '</option>';
    }).join('');
    return '<select class="cell' + (v ? '' : ' muted') + '" data-change="line" ' + ids(k) + ' data-field="' + field + '" data-key="' + esc(keyOf(k, field)) +
      '" aria-label="' + esc(label) + '">' + opts + '</select>';
  }
  function field(label, control) { return '<div class="tt-f"><span class="tt-fl">' + esc(label) + '</span><div class="tt-fc">' + control + '</div></div>'; }
  function proofBlock(k, line) {
    var own = { customerId: k.customerId, lineId: k.lineId, id: k.id, isExtra: false };   // paid how and paid belong to the line
    var paid = '<input class="cell num tt-paidin" type="text" inputmode="decimal" data-change="line" ' + ids(own) + ' data-field="paid" data-num="1" data-key="' +
      esc(keyOf(k, 'paid')) + '" value="' + esc(k.paid === null ? '' : fmt.m(k.paid)) + '" placeholder="0.0" aria-label="Paid so far, SEK m">' +
      '<span class="tt-of">of ' + esc(fmt.m(k.money)) + ' (SEK m)</span>';
    var sw = SWITCHES.map(function (s) {
      var on = !!k[s[0]];
      return '<button type="button" class="tt-switch" role="switch" aria-checked="' + on + '" data-action="trackSwitch" ' + ids(k) + ' data-field="' + s[0] +
        '" data-key="' + esc(keyOf(k, s[0])) + '"><span class="tt-knob" aria-hidden="true"></span><span>' + esc(s[1]) + '</span></button>';
    }).join('');
    return '<section class="ca-blk ca-b3"><h3 class="label">Proof and payment</h3><div class="tt-pp">' +
      field('Proof of execution', selectOf(k, 'proof', k.proof, 'Proof of execution', E.PROOFS, false)) +
      field('How it is paid', selectOf(own, 'paidHow', line ? line.paidHow : '', 'How it is paid', PAID_HOW, true)) +
      field('Paid so far', '<div class="tt-paidf">' + paid + '</div>') + '</div><div class="tt-sws">' + sw + '</div></section>';
  }
  // the open KPI (SPEC10 2): one row, one cell over every column, holding one card: the head (the KPI, its status and the chevron
  // on the table's columns, no month cells) and the detail (the three blocks)
  function card(c, k, scopeAll, span) {
    return '<tr class="detrow tt-detrow" data-kpi="' + esc(k.id) + '"><td colspan="' + span + '"><div class="tt-card">' +
      '<div class="tt-cardhead" data-action="trackDetails" data-href="' + esc(trackHash(k, scopeAll, true)) + '"><div class="tt-kpi">' + kpiText(k, scopeAll) +
      '</div><div></div><div class="tt-hst">' + ui.statusChip(k.status) + '</div><div class="tt-chcell">' + chevBtn(k, true, scopeAll) + '</div></div>' +
      '<div class="ca-det tt-det">' + agreedBlock(k) + monthsBlock(k) + proofBlock(k, lineOf(c, k)) + '</div></div></td></tr>';
  }
  function target(el) {
    var c = TTS.customer(el.getAttribute('data-cust'));
    return objOf(c, { lineId: el.getAttribute('data-line'), id: el.getAttribute('data-kpi'), isExtra: !!el.getAttribute('data-kpi') });
  }
  // a month cell of an ended or the current month: records the planned number (the target level for a %), a second click clears it
  TTS.actions.trackMonth = function (el) {
    var obj = target(el), m = +el.getAttribute('data-month');
    if (!obj || !(m >= 0 && m < 12)) return;
    if (!Array.isArray(obj.done)) obj.done = MON.map(function () { return null; });
    obj.done[m] = typeof obj.done[m] === 'number' ? null : Array.isArray(obj.plan) && typeof obj.plan[m] === 'number' ? obj.plan[m] : 0;
    TTS.focusKey = el.getAttribute('data-key');
    TTS.commit();
  };
  TTS.actions.trackEdit = function (el) {
    var id = el.getAttribute('data-id');
    editing = editing === id ? null : id;
    TTS.focusKey = el.getAttribute('data-key');
    TTS.render();
  };
  TTS.actions.trackSwitch = function (el) {
    var obj = target(el), f = el.getAttribute('data-field');
    if (!obj || SWITCHES.every(function (s) { return s[0] !== f; })) return;
    obj[f] = obj[f] !== true;
    TTS.focusKey = el.getAttribute('data-key');
    TTS.commit();
  };

  // ------------------------------------------------------------ views
  function viewOne(c, r) {
    var kpis = E.kpis(c);
    if (!kpis.length) return emptyKpis(c.id);
    var openId = openFrom(r, c.id);
    if (openId === 'none') openId = '';   // #track/<c> and #track/<c>/none: every panel closed (SPEC9 3)
    var cs = cols();
    var rows = kpis.map(function (k) {
      return k.id === openId ? card(c, k, false, cs.length) : closedRow(k, false);
    }).join('');
    return band(kpis, [c]) + legend() +
      '<div class="tt-wrap"><table class="grid tt-grid one">' + colgroup(cs) + thead(cs) + '<tbody>' + rows + '</tbody></table></div>' +
      ui.bottomRow(noKpiNote(c));
  }
  function viewAll(r) {
    var kpis = E.allKpis();
    if (!kpis.length) return emptyKpis('all');
    var openC = r && r.option, openK = r && r.sub;
    var custOrder = {};
    TTS.state.customers.forEach(function (c, i) { custOrder[c.id] = i; });
    var sorted = kpis.map(function (k, i) { return { k: k, i: i }; }).sort(function (a, b) {
      return (ORDER[a.k.status] - ORDER[b.k.status]) || (b.k.money - a.k.money) || (custOrder[a.k.customerId] - custOrder[b.k.customerId]) || (a.i - b.i);
    }).map(function (x) { return x.k; });
    var cs = cols();
    var rows = sorted.map(function (k) {
      return k.customerId === openC && k.id === openK ? card(TTS.customer(k.customerId), k, true, cs.length) : closedRow(k, true);
    }).join('');
    return band(kpis) + legend() +
      '<div class="tt-wrap"><table class="grid tt-grid all">' + colgroup(cs) + thead(cs) + '<tbody>' + rows + '</tbody></table></div>';
  }

  // ------------------------------------------------------------ demo date control (page head)
  function demoControl() {
    var st = TTS.state, d = validIso(st.demoDate) ? st.demoDate : '';
    return '<label class="tt-demo"><span class="label muted">Demo date</span>' +
      '<span class="tt-pill">' + ui.icon('calendar') + '<span class="tt-pilltext">' + esc(d ? fmt.date(d) : 'Pick a date') + '</span>' +
      '<input type="date" class="tt-dateinput" data-change="demoDate" data-action="demoPick" data-key="track-demo-date" value="' + esc(d) +
      '" min="' + st.year + '-01-01" max="' + st.year + '-12-31" aria-label="Demo date"></span></label>';
  }
  TTS.changes.demoDate = function (el) {
    var v = el.value;
    if (!validIso(v)) { TTS.render(); return; }
    TTS.state.demoDate = v;
    TTS.commit();
  };
  TTS.actions.demoPick = function (el) {
    if (typeof el.showPicker === 'function') { try { el.showPicker(); } catch (e) { /* picker not allowed here; the field still works by keyboard */ } }
  };

  TTS.screens.track = {
    head: function (scope, customer) {
      return {
        title: customer ? '<span class="hl">Track &amp; trace</span>: the KPIs to follow' : '<span class="hl">Track &amp; trace</span>, all customers',
        subtitle: 'The year for each KPI. Click a KPI to see what was agreed, the months and the proof.',
        actions: demoControl()
      };
    },
    one: function (c, r) { scrollOnOpen(r); return viewOne(c, r); },
    all: function (r) { scrollOnOpen(r); return viewAll(r); }
  };
})();

// ====================================================================================================
/* Screen 5 — Negotiation plan. Registers into TTS.screens.plan (SPEC 8.5). All class names start with "pl-". */
(function () {
  'use strict';
  var E = TTS.engine, fmt = TTS.fmt, ui = TTS.ui, esc = ui.esc, icon = ui.icon, hashFor = TTS.hashFor, listJoin = TTS.listJoin;
  var EDIT_KEYS = ['opening', 'propose', 'ask', 'give', 'obj1q', 'obj1a', 'obj2q', 'obj2a', 'obj3q', 'obj3a', 'obj4q', 'obj4a'];
  var SECTIONS = [['opening', 'Opening'], ['propose', 'What we propose']];
  var ASK_GIVE = 'What we ask and what we give', ASK_GIVE_COLS = [['ask', 'We ask'], ['give', 'We give']];

  function S() { return TTS.state; }
  function r1(n) { return Math.round((+n || 0) * 10 + 1e-9) / 10; }
  function lcFirst(s) { s = String(s || ''); return s.charAt(0).toLowerCase() + s.slice(1); }
  // period label in lower case, but a month name keeps its capital ("from March", "Feb to Nov")
  function periodLower(start, end) {
    var p = fmt.period(start, end);
    return /^[A-Z][a-z]{2} to /.test(p) ? p : lcFirst(p);
  }
  function hasEdits(c) { return !!c && Object.keys(c.scriptEdits || {}).some(function (k) { return EDIT_KEYS.indexOf(k) >= 0; }); }
  function norm(t) {
    return String(t || '').replace(/\r/g, '').replace(/ /g, ' ').replace(/[ \t]+\n/g, '\n').replace(/\s+$/, '');
  }
  function entryFor(list, id) { for (var i = 0; i < list.length; i++) if (list[i].priorityId === id) return list[i]; return { amount: 0, share: 0 }; }

  // ---------------------------------------------------------------- generated text (SPEC 8.5, fixed text blocks)
  function generate(c) {
    var st = S();
    var mwId = E.mustWin(c), mw = E.priority(mwId) || { name: '', winning: '' };
    var sT = E.summary(c.lines, 'today'), sP = E.summary(c.lines, 'proposed');
    var mwT = entryFor(E.spendByPriority(c.lines, 'today'), mwId).amount;
    var mwShare = sT.total > 0 ? mwT / sT.total : 0;
    var mwP = entryFor(E.spendByPriority(c.lines, 'proposed'), mwId).amount;
    var delta = mwP - mwT, up = r1(delta) > 0;
    var t = {};

    // Opening
    var a = r1(mwT) === 0 ? 'None of it goes to this priority.' : fmt.money(mwT) + ' of it (' + fmt.pct(mwShare) + ') goes to this priority.';
    var q = Math.round((1 - sT.shareConditional) * 100);
    var b = q === 0 ? 'Every line already has a counterpart with a number.' : q + '% of the total has no counterpart with a number.';
    t.opening = ('We want to invest more with you behind one priority for ' + st.year + ': ' + mw.name + '. ' + (mw.winning || '')).trim() +
      '\n\nToday we invest ' + fmt.money(sT.total) + ' with ' + c.name + '. ' + a + ' ' + b;

    // What we propose
    var p = [];
    p.push(up ? 'Move ' + fmt.money(delta) + ' into ' + mw.name + ': ' + fmt.money(mwP) + ', against ' + fmt.money(mwT) + ' today.'
      : r1(mwP) === 0 && r1(mwT) === 0 ? 'Agree a first investment line for ' + mw.name + ': the proposal has none yet.'
        : 'Keep ' + fmt.money(mwP) + ' on ' + mw.name + ' and agree a number for each line.');
    p.push('Link ' + fmt.pct(sP.shareLinked) + ' of the investment to our shared priorities, against ' + fmt.pct(sT.shareLinked) + ' today.');
    var nn = c.lines.filter(function (l) { return E.classOf(l, 'today') === 'partial'; });
    var p3 = 'Give every conditional line a counterpart, a number and a date, so both sides can follow it during the year.';
    if (nn.length) {
      var nnSum = nn.reduce(function (s, l) { return s + E.amount(l, 'today'); }, 0);
      p3 += ' Today ' + nn.length + (nn.length === 1 ? ' line' : ' lines') + ' worth ' + fmt.money(nnSum) + (nn.length === 1 ? ' is' : ' are') +
        ' only partially conditional: ' + listJoin(nn.map(function (l) { return l.name; })) + '.';
    }
    p.push(p3);
    p.push(r1(sT.total) === r1(sP.total) ? 'The total stays at ' + fmt.money(sT.total) + '.'
      : 'The total moves from ' + fmt.money(sT.total) + ' to ' + fmt.money(sP.total) + '.');
    t.propose = p.map(function (x) { return '– ' + x; }).join('\n');

    // What we ask and what we give
    var cond = c.lines.filter(function (l) { return E.lineStatus(l, 'proposed') === 'conditional'; })
      .map(function (l, i) { return { l: l, i: i }; })
      .sort(function (x, y) { return (E.amount(y.l, 'proposed') - E.amount(x.l, 'proposed')) || (x.i - y.i); })
      .map(function (x) { return x.l; });
    var ask = cond.slice(0, 4).map(function (l) {
      return '– ' + String(l.counterpart).trim() + ': ' + fmt.target(l.target, l.unit) + ', ' + periodLower(l.start, l.end);
    });
    c.lines.filter(function (l) { return E.classOf(l, 'proposed') === 'partial'; }).slice(0, 2).forEach(function (l) {
      ask.push('– A number and a date for ' + l.name + '.');
    });
    var give = ['– ' + fmt.money(sP.conditional) + ' linked to these counterparts, paid when they are delivered.'];
    if (up) give.push('– ' + fmt.money(delta) + ' more for ' + mw.name + '.');
    give.push('– One KPI list with an owner and a data source for each line, reviewed together every quarter.');
    t.ask = ask.join('\n');
    t.give = give.join('\n');

    // Objections
    var d = sT.unconditional - sP.unconditional;
    t.obj1q = 'We need the unconditional discount to stay as it is.';
    t.obj1a = r1(sP.unconditional) < r1(sT.unconditional)
      ? 'The total does not go down. We move ' + fmt.money(d) + ' from the unconditional part to lines you can earn in full by delivering the counterparts.'
      : 'The unconditional part stays at ' + fmt.money(sT.unconditional) + '. What changes is that every other line gets a number and a date.';
    t.obj2q = 'These targets are too hard to reach.';
    t.obj2a = 'Each target is something you control in your own business: listings, shelf space, displays and promotion execution. ' +
      'None of them depends on how much you buy from us. We start from what you deliver today and review the numbers together every quarter.';
    var seen = {}, sources = [];
    cond.forEach(function (l) {
      var s = String(l.source || '').trim(), k = s.toLowerCase();
      if (s && !seen[k] && sources.length < 3) { seen[k] = 1; sources.push(lcFirst(s)); }
    });
    t.obj3q = 'We cannot share that data.';
    t.obj3a = 'We only need what shows that the counterpart was delivered: ' + (sources.length ? listJoin(sources) : 'a simple report agreed with you') +
      '. We do not ask for data on other suppliers.';
    var kp = E.kpis(c);
    t.obj4q = 'What happens if we miss a target?';
    t.obj4a = 'Only the money linked to that KPI is at stake, and we look at it together before the period ends. Today ' +
      fmt.money(E.moneyAtStake(kp)) + ' depends on ' + kp.length + (kp.length === 1 ? ' KPI' : ' KPIs') + '; the rest of the agreement is not affected.';
    return t;
  }
  function texts(c) {
    var g = generate(c), ed = c.scriptEdits || {}, out = {};
    EDIT_KEYS.forEach(function (k) { out[k] = Object.prototype.hasOwnProperty.call(ed, k) ? String(ed[k]) : g[k]; });
    return out;
  }

  // ---------------------------------------------------------------- card 1: spend against priorities
  function chart(lines) {
    var today = E.spendByPriority(lines, 'today'), prop = E.spendByPriority(lines, 'proposed');
    var max = 0;
    today.concat(prop).forEach(function (e) { if (e.share > max) max = e.share; });
    function bar(share, cls) {
      var w = max > 0 ? share / max * 100 : 0;
      return '<div class="pl-bar ' + cls + (share > 0 ? '' : ' zero') + '" style="width:' + (share > 0 ? w.toFixed(2) + '%' : '6px') + '"></div>';
    }
    return '<div class="pl-chart">' + today.map(function (e, i) {
      var pe = prop[i];
      var tp = Math.round(e.share * 100), pp = Math.round(pe.share * 100);
      return '<div class="pl-pri' + (e.priorityId ? '' : ' none') + '">' +
        '<div class="pl-pri-top"><span class="pl-pri-name">' + esc(e.name) + '</span>' +
        '<span class="pl-pri-pct">' + tp + '% <span class="pl-arrow">→</span> <span class="' + (pp > tp ? 'up' : '') + '">' + pp + '%</span></span></div>' +
        '<div class="pl-bars">' + bar(e.share, 'today') + bar(pe.share, 'prop') + '</div>' +
        '<div class="pl-pri-sub">SEK ' + fmt.m(e.amount) + ' m → ' + fmt.m(pe.amount) + ' m</div></div>';
    }).join('') + '</div>';
  }
  function legend(items) {
    return '<div class="pl-legend">' + items.map(function (x) { return '<span><i class="dot ' + x[0] + '"></i>' + esc(x[1]) + '</span>'; }).join('') + '</div>';
  }
  function cardTop(label, right) {
    return '<div class="pl-top"><span class="label">' + esc(label) + '</span>' + (right || '') + '</div>';
  }
  function mustWinBlock(c) {
    var sug = E.suggestedMustWin(c), mw = E.mustWin(c);
    var opts = S().priorities.map(function (p) {
      return '<option value="' + esc(p.id) + '"' + (p.id === mw ? ' selected' : '') + '>' + esc(p.name) + '</option>';
    }).join('');
    var note;
    if (mw === sug) {
      var amt = entryFor(E.spendByPriority(c.lines, 'today'), sug).amount;
      note = '<span>' + (r1(amt) === 0 ? 'Suggested: no spend today.' : 'Suggested: lowest spend today, ' + esc(fmt.money(amt)) + '.') + '</span>';
    } else {
      var sp = E.priority(sug);
      note = '<span>Your choice. Suggested: ' + esc(sp ? sp.name : '') + '.</span>' +
        '<button type="button" class="linkbtn pl-use" data-action="planUseSuggestion" data-cust="' + esc(c.id) + '" data-key="plan:' + esc(c.id) + ':usesug">Use the suggestion</button>';
    }
    var mwp = E.priority(mw);
    return '<div class="pl-mw"><label class="label" for="pl-mw-sel">Must-win battle</label>' +
      '<div class="pl-mwsel"><select id="pl-mw-sel" data-change="planMustWin" data-cust="' + esc(c.id) + '" data-key="plan:' + esc(c.id) + ':mw">' + opts + '</select>' +
      icon('caret-down', 18) + '</div>' +
      '<div class="pl-mwprint">' + esc(mwp ? mwp.name : '') + '</div>' +
      '<div class="pl-mwnote">' + note + '</div></div>';
  }

  // ---------------------------------------------------------------- card 2: opportunities
  function opportunities(c) {
    var items = E.opportunities(c);
    if (!items.length) return '<p class="pl-none">Nothing stands out. Every line is conditional, has a number and serves a priority.</p>';
    return '<ol class="pl-opps">' + items.map(function (o, i) {
      return '<li><span class="pl-n">' + (i + 1) + '</span><div><div class="pl-ot">' + esc(o.title) + '</div><div class="pl-od">' + esc(o.detail) + '</div></div></li>';
    }).join('') + '</ol>';
  }

  function oppSize(c) {
    var total = E.summary(c.lines, 'today').total, size = E.opportunitySize(c);
    return '<div class="pl-oppsum stat"><div class="v"><span class="u pre">SEK</span>' + esc(fmt.m(size)) + '<span class="u">m</span></div>' +
      '<div class="l">' + fmt.pct(total > 0 ? size / total : 0) + ' of spend is not yet conditional with a number and linked to a priority</div></div>';
  }

  // ---------------------------------------------------------------- card 3: script
  function editable(c, key, text, label, cls) {
    return '<div class="pl-edit' + (cls ? ' ' + cls : '') + '" contenteditable="plaintext-only" spellcheck="false" role="textbox" aria-multiline="true" aria-label="' + esc(label) +
      '" data-plan-edit="' + key + '" data-cust="' + esc(c.id) + '" data-key="plan:' + esc(c.id) + ':' + key + '">' + esc(text) + '</div>';
  }
  function scriptCard(c, tab) {
    var t = texts(c);
    var obj = tab === 'objections';
    var seg = '<div class="segctl" role="tablist" aria-label="Script or objections">' +
      '<a role="tab" class="' + (obj ? '' : 'on') + '" aria-selected="' + !obj + '" href="' + hashFor('plan', c.id, '') + '" data-key="plan-tab-script">Script</a>' +
      '<a role="tab" class="' + (obj ? 'on' : '') + '" aria-selected="' + obj + '" href="' + hashFor('plan', c.id, 'objections') + '" data-key="plan-tab-obj">Objections</a></div>';
    var rebuild = '<button type="button" class="btn ghost pl-rebuild" data-action="planRebuild" data-cust="' + esc(c.id) + '" data-key="plan:' + esc(c.id) + ':rebuild"' +
      (hasEdits(c) ? '' : ' hidden') + '>' + icon('reset') + 'Rebuild from the data</button>';
    var script = '<div class="pl-pane pl-script' + (obj ? ' pl-off' : '') + '" role="tabpanel" aria-label="Script"><div class="pl-printlabel">Script</div>' + SECTIONS.map(function (s) {
      return '<section class="pl-sec"><div class="label">' + esc(s[1]) + '</div>' + editable(c, s[0], t[s[0]], s[1], s[0] === 'opening' ? '' : 'pl-list') + '</section>';
    }).join('') + '<section class="pl-sec"><div class="label">' + esc(ASK_GIVE) + '</div><div class="pl-cols">' + ASK_GIVE_COLS.map(function (k) {
      return '<div class="pl-col"><div class="pl-colh">' + esc(k[1]) + '</div>' + editable(c, k[0], t[k[0]], k[1], 'pl-list') + '</div>';
    }).join('') + '</div></section></div>';
    var objs = '<div class="pl-pane pl-objs' + (obj ? '' : ' pl-off') + '" role="tabpanel" aria-label="Objections"><div class="pl-printlabel">Objections</div>' + [1, 2, 3, 4].map(function (n) {
      return '<section class="pl-obj">' + editable(c, 'obj' + n + 'q', t['obj' + n + 'q'], 'Objection ' + n, 'pl-q') +
        editable(c, 'obj' + n + 'a', t['obj' + n + 'a'], 'Answer to objection ' + n, 'pl-a') + '</section>';
    }).join('') + '</div>';
    return '<div class="card pl-card pl-scriptcard"><div class="pl-top">' + seg + rebuild + '</div><div class="pl-scroll">' + script + objs + '</div></div>';
  }

  // ---------------------------------------------------------------- key figures (SPEC5 8): five groups, each a link to its screen
  function kfData(cs) {
    var kf = E.keyFigures(cs), s = kf.sales, cl = kf.classes, cp = kf.counterparts, kp = kf.kpis;
    var sh = function (x) { return Math.round(x.share * 100) + '%'; };
    return { kf: kf, third: s.thirdNet, thirdPct: s.gross > 0 ? s.thirdNet / s.gross : 0, u: sh(cl.unconditional), p: sh(cl.partial), c: sh(cl.conditional), cp: cp, kp: kp };
  }
  function kfBand(cs, scope) {
    var d = kfData(cs), kf = d.kf, one = scope !== 'all';
    var href = {
      terms: hashFor('agreement', scope, one ? 'waterfall' : ''), cls: hashFor('agreement', scope, ''), cps: hashFor('agreement', scope, ''),
      kpis: hashFor('track', scope, ''), ahead: hashFor('track', scope, '')
    };
    function grp(key, label, body) {
      return '<a class="pl-kf pl-kf-' + key + '" href="' + href[key] + '" data-key="plan-kf:' + key + '"><span class="label">' + esc(label) +
        '<span class="pl-kfgo">' + icon('arrow-right', 14) + '</span></span>' + body + '</a>';
    }
    var cl = kf.classes, tot = cl.unconditional.amount + cl.partial.amount + cl.conditional.amount;
    var seg = function (x, cls) { return x.amount > 1e-9 ? '<i class="' + cls + '" style="flex-grow:' + (tot > 0 ? x.amount / tot : 0).toFixed(4) + '"></i>' : ''; };
    var chips = [['executed', 'green', 'Fully executed'], ['partly', 'yellow', 'Partially executed'], ['notexecuted', 'outline', 'Not executed']].map(function (x) {
      return '<span class="chip ' + x[1] + (d.kp[x[0]] ? '' : ' zero') + '"><b>' + d.kp[x[0]] + '</b>' + x[2] + '</span>';
    }).join('');
    var aheadZero = fmt.m(kf.paidAhead) === '0.0';
    return '<nav class="pl-kfband" aria-label="Key figures">' +
      grp('terms', 'Trade terms', '<div class="stat"><div class="v"><span class="u pre">SEK</span>' + esc(fmt.m(kf.total)) + '<span class="u">m</span></div></div>' +
        '<div class="pl-kf1">' + esc(fmt.pct1(kf.pctOfGross)) + ' of gross revenue</div>' +
        '<div class="pl-kf2">Third net ' + esc(fmt.money(d.third)) + ' · ' + esc(fmt.pct1(d.thirdPct)) + '</div>') +
      grp('cls', 'Classification', '<div class="pl-kfbar" role="img" aria-label="' + esc(d.u + ' unconditional, ' + d.p + ' partially conditional, ' + d.c + ' conditional') + '">' +
        seg(cl.unconditional, 'u') + seg(cl.partial, 'n') + seg(cl.conditional, 'c') + '</div>' +
        '<div class="pl-kf1 pl-kfcls"><span><i class="dot u"></i><b>' + d.u + '</b> unconditional</span><span><i class="dot n"></i><b>' + d.p + '</b> partially</span>' +
        '<span><i class="dot c"></i><b>' + d.c + '</b> conditional</span></div>') +
      grp('cps', 'Counterparts', '<div class="pl-kf1"><b>' + d.cp.smart + ' of ' + d.cp.lines + '</b> SMART</div>' +
        '<div class="pl-kf1"><b>' + d.cp.aligned + ' of ' + d.cp.lines + '</b> aligned with a priority</div>' +
        (r1(d.cp.smartNotAlignedAmount) > 0 ? '<div class="pl-kf2">' + esc(fmt.money(d.cp.smartNotAlignedAmount)) + ' SMART but not aligned</div>' : '')) +
      grp('kpis', 'KPIs', '<div class="pl-kfchips">' + chips + '</div>') +
      grp('ahead', 'Paid ahead of delivery', '<div class="stat ' + (aheadZero ? 'pl-kfnone' : 'pl-kfahead') + '"><div class="v"><span class="u pre">SEK</span>' + esc(fmt.m(kf.paidAhead)) + '<span class="u">m</span></div></div>') +
      '</nav>';
  }
  // the same figures as one compact line for the one-page print (shown only in the print layout)
  function kfPrint(c) {
    var d = kfData([c]), kf = d.kf, kp = d.kp;
    var parts = [
      ['Trade terms', fmt.money(kf.total) + ', ' + fmt.pct1(kf.pctOfGross) + ' of gross'],
      ['', d.u + ' unconditional, ' + d.p + ' partially, ' + d.c + ' conditional'],
      ['', d.cp.smart + ' of ' + d.cp.lines + ' SMART, ' + d.cp.aligned + ' of ' + d.cp.lines + ' aligned'],
      ['KPIs', kp.executed + ' fully executed, ' + kp.partly + ' partially executed, ' + kp.notexecuted + ' not executed'],
      ['Paid ahead', fmt.money(kf.paidAhead)]
    ];
    return '<div class="pl-kfprint">' + parts.map(function (p) { return '<span>' + (p[0] ? '<b>' + esc(p[0]) + '</b> ' : '') + esc(p[1]) + '</span>'; }).join('') + '</div>';
  }

  function printHead(title) {
    var ex = S().customers.some(function (x) { return x.example; });
    return '<div class="pl-printhead"><span>' + esc(title) + '</span>' + (ex ? '<span class="badge">Example data</span>' : '') + '</div>';
  }

  function planOne(c, r) {
    var tab = r && r.option === 'objections' ? 'objections' : 'script';
    return '<div class="pl-view pl-one">' + printHead('Negotiation plan — ' + c.name) + kfPrint(c) + kfBand([c], c.id) + '<div class="pl-grid">' +
      '<div class="card pl-card pl-spend">' + cardTop('Spend against priorities', legend([['grey', 'Today'], ['green', 'Proposed']])) +
      chart(c.lines) + mustWinBlock(c) + '</div>' +
      '<div class="card pl-card pl-oppcard">' + cardTop('Top opportunities') + opportunities(c) + oppSize(c) + '</div>' +
      scriptCard(c, tab) + '</div></div>';
  }

  // ---------------------------------------------------------------- all customers
  function rankedCard() {
    var cs = S().customers.map(function (c) {
      var total = E.summary(c.lines, 'today').total, size = E.opportunitySize(c);
      return { c: c, size: size, share: total > 0 ? size / total : 0 };
    }).sort(function (a, b) { return b.size - a.size; });
    var max = 0;
    cs.forEach(function (x) { if (x.size > max) max = x.size; });
    return '<div class="card pl-card">' + cardTop('Customers ranked by opportunity') +
      '<p class="pl-caption">Spend that is not yet conditional with a number and linked to a priority.</p><div class="pl-rows">' +
      cs.map(function (x) {
        var mw = E.priority(E.mustWin(x.c));
        return '<a class="pl-crow" href="' + hashFor('plan', x.c.id, '') + '" data-key="plan-open:' + esc(x.c.id) + '">' +
          '<div class="pl-crow-top"><span class="pl-cname">' + esc(x.c.name) + '</span><span class="pl-cval"><b>' + esc(fmt.money(x.size)) + '</b>' +
          '<span class="pl-cshare">' + fmt.pct(x.share) + ' of spend</span></span></div>' +
          '<div class="pl-track"><i style="width:' + (max > 0 ? (x.size / max * 100).toFixed(2) : 0) + '%"></i></div>' +
          '<div class="pl-crow-sub">Must-win battle: ' + esc(mw ? mw.name : '–') + '</div></a>';
      }).join('') + '</div></div>';
  }
  function growthCard() {
    var cs = S().customers.map(function (c) { return { c: c, g: E.growth(c) }; });
    var pos = 0, neg = 0;
    cs.forEach(function (x) {
      [x.g.investment, x.g.revenue].forEach(function (v) { if (v > pos) pos = v; if (-v > neg) neg = -v; });
    });
    var span = pos + neg || 1, zero = neg / span * 100;
    // the scale runs inside the plot with 60 px kept free at each side for the value labels (so a label never leaves its box)
    function at(pc) { return 'calc(60px + (100% - 120px) * ' + (pc / 100).toFixed(4) + ')'; }
    function gbar(v, cls) {
      var w = Math.abs(v) / span * 100;
      var left = v >= 0 ? zero : zero - w;
      var lab = v >= 0 ? 'left:calc(' + at(zero + w) + ' + 8px)' : 'right:calc(' + at(100 - (zero - w)) + ' + 8px)';
      return '<div class="pl-gline"><i class="' + cls + '" style="left:' + at(left) + ';width:' + (w > 0 ? 'calc((100% - 120px) * ' + (w / 100).toFixed(4) + ')' : '2px') + '"></i>' +
        '<span class="pl-gval" style="' + lab + '">' + esc(fmt.pct1s(v)) + '</span></div>';
    }
    return '<div class="card pl-card">' + cardTop('Investment growth against revenue growth') +
      '<div class="pl-caption">' + legend([['navy', 'Investment'], ['blue', 'Revenue']]) + '</div><div class="pl-rows">' + cs.map(function (x) {
        return '<div class="pl-grow"><div class="pl-crow-top"><span class="pl-cname">' + esc(x.c.name) + '</span>' +
          (x.g.faster ? ui.chip('yellow', 'Faster than revenue') : '') + '</div>' +
          '<div class="pl-gplot"><span class="pl-zero" style="left:' + at(zero) + '"></span>' +
          gbar(x.g.investment, 'inv') + gbar(x.g.revenue, 'rev') + '</div></div>';
      }).join('') + '</div></div>';
  }
  // one headline above the three cards (SPEC2 7): "SEK 67.2 m of 95.0 m is not yet …"
  function headline() {
    var size = 0, total = 0;
    S().customers.forEach(function (c) { size += E.opportunitySize(c); total += E.summary(c.lines, 'today').total; });
    return '<div class="pl-headband"><span class="stat"><span class="v"><span class="u pre">SEK</span>' + esc(fmt.m(size)) + '<span class="u">m</span></span></span>' +
      '<span class="pl-headtext">of ' + esc(fmt.m(total)) + ' m is not yet conditional with a number and linked to a priority</span></div>';
  }
  function planAll() {
    return '<div class="pl-view pl-all">' + headline() + kfBand(S().customers, 'all') + '<div class="pl-grid pl-grid3">' +
      '<div class="card pl-card pl-spend">' + cardTop('Spend against priorities', legend([['grey', 'Today'], ['green', 'Proposed']])) +
      chart(TTS.scopeLines('all')) + '</div>' + rankedCard() + growthCard() + '</div></div>';
  }

  // ---------------------------------------------------------------- copy text
  function plainText(c) {
    var t = texts(c);
    var out = ['Negotiation plan — ' + c.name, ''];
    SECTIONS.forEach(function (s) { out.push(s[1], t[s[0]], ''); });
    out.push(ASK_GIVE, 'We ask:', t.ask, '', 'We give:', t.give, '');
    out.push('Objections', '');
    [1, 2, 3, 4].forEach(function (n) { out.push(t['obj' + n + 'q'], t['obj' + n + 'a'], ''); });
    return out.join('\n').replace(/\n+$/, '\n');
  }
  function fallbackCopy(text) {
    var ta = document.createElement('textarea');
    ta.value = text;
    ta.setAttribute('readonly', '');
    ta.style.position = 'fixed'; ta.style.top = '-1000px'; ta.style.left = '0'; ta.style.opacity = '0';
    document.body.appendChild(ta);
    ta.select();
    var ok = false;
    try { ok = document.execCommand('copy'); } catch (e) { ok = false; }
    document.body.removeChild(ta);
    return ok;
  }
  function copyText(text) {
    function done(ok) { ui.toast(ok ? 'Script copied.' : 'The script could not be copied.'); }
    var cb = typeof navigator !== 'undefined' && navigator.clipboard;
    if (cb && typeof cb.writeText === 'function') {
      try {
        cb.writeText(text).then(function () { done(true); }, function () { done(fallbackCopy(text)); });
        return;
      } catch (e) { /* fall through */ }
    }
    done(fallbackCopy(text));
  }

  // ---------------------------------------------------------------- screen registration
  TTS.screens.plan = {
    head: function (scope, c) {
      if (!c) return { title: '<span class="hl">Negotiation plan</span>, all customers', subtitle: 'Generated from the agreement: where the money goes, what to ask and what to say.' };
      return {
        title: '<span class="hl">Negotiation plan</span> for ' + esc(c.name),
        subtitle: 'Generated from the agreement: where the money goes, what to ask and what to say.',
        actions: '<button type="button" class="btn secondary" data-action="planCopy" data-cust="' + esc(c.id) + '" data-key="plan-copy">' + icon('copy') + 'Copy script</button>' +
          '<button type="button" class="btn primary" data-action="planPrint" data-key="plan-print">' + icon('print') + 'Print</button>'
      };
    },
    one: function (c, r) { return planOne(c, r); },
    all: function () { return planAll(); },
    afterRender: function (root) {
      fitOppSum(root);
      // the web font may still be loading on first paint: measure again once it is in
      if (document.fonts && document.fonts.ready) document.fonts.ready.then(function () { fitOppSum(document.getElementById('view')); });
    }
  };
  // The opportunity total sits at the foot of card 2 only when card 2 then stays no taller than the cards beside it
  // (it never makes the row taller). Content-based, not tied to the window height.
  function fitOppSum(root) {
    var card = root && root.querySelector('.pl-oppcard'), sum = card && card.querySelector('.pl-oppsum');
    if (!sum) return;
    var grid = card.parentNode;
    sum.classList.remove('pl-nofit');
    grid.style.alignItems = 'start';
    var others = 0;
    Array.prototype.forEach.call(grid.children, function (x) { if (x !== card && x.offsetTop === card.offsetTop) others = Math.max(others, x.offsetHeight); });
    var mine = card.offsetHeight;
    grid.style.alignItems = '';
    if (mine > others) sum.classList.add('pl-nofit');
  }
  TTS.plan ={ generate: generate, texts: texts, plainText: plainText };

  TTS.changes.planMustWin = function (el) {
    var c = TTS.customer(el.getAttribute('data-cust'));
    if (!c) return;
    var v = el.value || null;
    c.mustWinId = (!v || v === E.suggestedMustWin(c)) ? null : v;
    TTS.commit();
  };
  TTS.actions.planUseSuggestion = function (el) {
    var c = TTS.customer(el.getAttribute('data-cust'));
    if (!c) return;
    c.mustWinId = null;
    TTS.focusKey = 'plan:' + c.id + ':mw';
    TTS.commit();
  };
  TTS.actions.planRebuild = function (el) {
    var c = TTS.customer(el.getAttribute('data-cust'));
    if (!c) return;
    ui.confirm({ title: 'Rebuild the script from the data?', text: 'Your edits to the script and the objections are replaced by the text generated from the agreement.', okLabel: 'Rebuild' })
      .then(function (ok) {
        if (!ok) return;
        c.scriptEdits = {};
        TTS.commit();
        ui.toast('Script rebuilt from the data.');
      });
  };
  TTS.actions.planCopy = function (el) {
    var c = TTS.customer(el.getAttribute('data-cust'));
    if (c) copyText(plainText(c));
  };
  TTS.actions.planPrint = function () { if (typeof window !== 'undefined' && window.print) window.print(); };

  // Edits are saved on blur without re-rendering, so the click that caused the blur still reaches its target.
  function saveEdit(el) {
    var c = TTS.customer(el.getAttribute('data-cust')), key = el.getAttribute('data-plan-edit');
    if (!c || EDIT_KEYS.indexOf(key) < 0) return;
    var text = norm(el.innerText !== undefined ? el.innerText : el.textContent);
    var gen = norm(generate(c)[key]);
    c.scriptEdits = c.scriptEdits || {};
    var had = Object.prototype.hasOwnProperty.call(c.scriptEdits, key);
    if (text !== gen) c.scriptEdits[key] = text;
    else if (had) delete c.scriptEdits[key];
    else return;
    TTS.save();
    var rb = document.querySelector('.pl-rebuild[data-cust="' + c.id.replace(/"/g, '\\"') + '"]');
    if (rb) rb.hidden = !hasEdits(c);
  }
  // Print: switch .pl-one to the print layout (.pl-pm) and, if it is taller than one A4 landscape page
  // (277 x 190 mm inside 10 mm margins = 1047 x 718 CSS px), scale it down with zoom until it fits.
  var PAGE_W = 1040, PAGE_H = 712;
  function beforePrint() {
    var el = document.querySelector('.pl-one');
    if (!el) return;
    el.classList.add('pl-pm');
    el.style.zoom = '';
    var z = 1;
    el.style.width = PAGE_W + 'px';
    while (z > 0.6 && el.offsetHeight * z > PAGE_H) {
      z = Math.round((z - 0.02) * 100) / 100;
      el.style.width = (PAGE_W / z).toFixed(1) + 'px';
    }
    if (z < 1) el.style.zoom = String(z);
  }
  function afterPrint() {
    var el = document.querySelector('.pl-one');
    if (!el) return;
    el.classList.remove('pl-pm');
    el.style.width = ''; el.style.zoom = '';
  }
  if (typeof document !== 'undefined') {
    document.addEventListener('focusout', function (e) {
      var el = e.target && e.target.closest && e.target.closest('[data-plan-edit]');
      if (el) saveEdit(el);
    });
    window.addEventListener('resize', function () { fitOppSum(document.getElementById('view')); });
    window.addEventListener('beforeprint', beforePrint);
    window.addEventListener('afterprint', afterPrint);
  }
})();

if (typeof module !== 'undefined' && module.exports) module.exports = TTS;
