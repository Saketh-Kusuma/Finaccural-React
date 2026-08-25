/**
 * Excel JS API layer — workbook, worksheet, range and formatting operations.
 *
 * Row writes are delegated to batchDataLoader.writeRowsInBatches() so a large
 * master-data set never lands in a single Excel.run() context.
 */
import { writeRowsInBatches } from "../batchDataLoader.js";
import {
    isChangedRecord,
    partitionExistingThenNew,
    countAllMasterDataRecords,
    flattenChangedMasterDataRecords
} from "./excelDataMappers.js";

const ExcelService = {
    /**
     * Populates Excel workbook sheet "1.Master_Data" with ERP payload.
     * @param {string} provider - Active ERP provider
     * @param {object} data - Master data payload
     */
    async writeMasterData(provider, data) {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItem("1.Master_Data");

            // Clear existing records in spreadsheet grid below headers.
            // "All" (not just "Contents") is required so any "new record"
            // highlight fill from a previous refresh doesn't linger on
            // rows that are no longer new — font size/wrap and column
            // width are reapplied unconditionally below anyway.
            const clearRange = sheet.getRange("A2:AB10000");
            clearRange.clear("All");

            // Group all data by Organization
            const orgGroupsMap = new Map();
            const fallbackOrgName = provider === "quickbooks" ? "QuickBooks Company" : "Xero Organisation";

            const getOrCreateGroup = (id, name) => {
                const cleanName = (name && name !== "Default" && name !== "Default Organization") ? name : fallbackOrgName;
                const key = cleanName.trim();

                if (!orgGroupsMap.has(key)) {
                    orgGroupsMap.set(key, {
                        id: cleanName,
                        name: cleanName,
                        accounts: [],
                        classes: [],
                        locations: [],
                        entities: []
                    });
                }
                return orgGroupsMap.get(key);
            };

            // 1. Group Companies
            const rawCompanies = Array.isArray(data.company)
                ? data.company
                : (data.company ? [data.company] : []);

            for (const c of rawCompanies) {
                if (c) getOrCreateGroup(c.name || c.id, c.name);
            }

            // 2. Group Accounts
            if (data.accounts) {
                for (const a of data.accounts) {
                    const nameKey = a.clientName || a.clientId || fallbackOrgName;
                    const group = getOrCreateGroup(a.clientId, nameKey);
                    group.accounts.push(a);
                }
            }

            // 3. Group Classes
            if (data.classes) {
                for (const c of data.classes) {
                    const nameKey = c.clientName || c.clientId || fallbackOrgName;
                    const group = getOrCreateGroup(c.clientId, nameKey);
                    group.classes.push(c);
                }
            }

            // 4. Group Locations
            if (data.locations) {
                for (const l of data.locations) {
                    const nameKey = l.clientName || l.clientId || fallbackOrgName;
                    const group = getOrCreateGroup(l.clientId, nameKey);
                    group.locations.push(l);
                }
            }

            // 5. Group Entities (Customers and Vendors)
            if (data.customers) {
                for (const cust of data.customers) {
                    const nameKey = cust.clientName || cust.clientId || fallbackOrgName;
                    const group = getOrCreateGroup(cust.clientId, nameKey);
                    group.entities.push({
                        clientId: nameKey,
                        name: cust.name || cust.DisplayName || cust.Name || "",
                        type: "Customer",
                        id: cust.id || cust.Id || cust.ContactID || "",
                        status: cust.active !== undefined ? (cust.active ? "Active" : "Inactive") : "Active",
                        isNew: !!cust.isNew,
                        isUpdated: !!cust.isUpdated
                    });
                }
            }

            if (data.vendors) {
                for (const vend of data.vendors) {
                    const nameKey = vend.clientName || vend.clientId || fallbackOrgName;
                    const group = getOrCreateGroup(vend.clientId, nameKey);
                    group.entities.push({
                        clientId: nameKey,
                        name: vend.name || vend.DisplayName || vend.Name || "",
                        type: "Vendor",
                        id: vend.id || vend.Id || vend.ContactID || "",
                        status: vend.active !== undefined ? (vend.active ? "Active" : "Inactive") : "Active",
                        isNew: !!vend.isNew,
                        isUpdated: !!vend.isUpdated
                    });
                }
            }

            // Write grouped data sequentially with 2 blank rows between organizations
            let currentRow = 2;

            for (const [key, group] of orgGroupsMap) {
                const orgName = group.name; // Client ID and Client Name are identical

                const accCount = group.accounts.length;
                const classCount = group.classes.length;
                const locCount = group.locations.length;
                const entityCount = group.entities.length;

                const maxRows = Math.max(1, accCount, classCount, locCount, entityCount);

                // Section 1 (A:B) - Client Config (Client ID and Client Name SAME)
                sheet.getRange(`A${currentRow}:B${currentRow}`).values = [[orgName, orgName]];

                // Section 2 (D:L) - Accounts, written in batches of
                // MASTER_DATA_WRITE_BATCH_SIZE rows so large chart-of-accounts
                // payloads don't sit in memory as one giant pending write.
                if (accCount > 0) {
                    // Unchanged accounts keep their original order and
                    // default background; new-or-updated ones are moved
                    // to the bottom and highlighted — never interleaved
                    // in the middle of the existing rows.
                    const orderedAccounts = partitionExistingThenNew(group.accounts);
                    const accValues = orderedAccounts.map(a => [
                        orgName,
                        a.acctNum || a.code || a.AcctNum || a.Code || "",
                        a.name || a.Name || "",
                        a.accountType || a.type || a.AccountType || a.Type || "",
                        a.accountSubType || a.description || a.AccountSubType || "",
                        a.classification || a.Classification || "",
                        a.fullyQualifiedName || a.name || a.Name || "",
                        a.active !== undefined ? (a.active ? "Active" : "Inactive") : "Active",
                        a.id || a.Id || a.AccountID || ""
                    ]);
                    const accHighlightFlags = orderedAccounts.map(isChangedRecord);
                    await writeRowsInBatches(context, sheet, "D", "L", currentRow, accValues, accHighlightFlags);
                }

                // Section 3 (N:Q) - Classes, batched the same way.
                if (classCount > 0) {
                    const orderedClasses = partitionExistingThenNew(group.classes);
                    const classValues = orderedClasses.map(c => [
                        orgName,
                        c.name || c.Name || "",
                        c.id || c.Id || "",
                        c.active !== undefined ? (c.active ? "Active" : "Inactive") : "Active"
                    ]);
                    const classHighlightFlags = orderedClasses.map(isChangedRecord);
                    await writeRowsInBatches(context, sheet, "N", "Q", currentRow, classValues, classHighlightFlags);
                }

                // Section 4 (S:V) - Locations, batched the same way.
                if (locCount > 0) {
                    const orderedLocations = partitionExistingThenNew(group.locations);
                    const locValues = orderedLocations.map(l => [
                        orgName,
                        l.name || l.Name || "",
                        l.id || l.Id || "",
                        l.active !== undefined ? (l.active ? "Active" : "Inactive") : "Active"
                    ]);
                    const locHighlightFlags = orderedLocations.map(isChangedRecord);
                    await writeRowsInBatches(context, sheet, "S", "V", currentRow, locValues, locHighlightFlags);
                }

                // Section 5 (X:AB) - Entities (Customers + Vendors), the
                // section most likely to be large, batched the same way.
                if (entityCount > 0) {
                    const orderedEntities = partitionExistingThenNew(group.entities);
                    const entityValues = orderedEntities.map(e => [
                        orgName,
                        e.name,
                        e.type,
                        e.id,
                        e.status
                    ]);
                    const entityHighlightFlags = orderedEntities.map(isChangedRecord);
                    await writeRowsInBatches(context, sheet, "X", "AB", currentRow, entityValues, entityHighlightFlags);
                }

                // Advance currentRow by maxRows + 2 (providing 2 empty row spaces between organizations!)
                currentRow += maxRows + 2;
            }

            // Apply consistent font size and wrap text to prevent overlapping/truncation
            const dataRange = sheet.getRange("A2:AB10000");
            dataRange.format.font.size = 11;
            dataRange.format.wrapText = true;

            sheet.getRange("A:AB").format.columnWidth = 115;

            await context.sync();
        });
    },

    /**
     * Appends only brand-new records from a Master Data Pull to the
     * "1.Master_Data" sheet — an incremental, append-only ledger
     * writer. NOT currently called by the UI: the Refresh Schedule
     * button now pulls and (re)writes the full master data set in
     * batches, the same as Pull Master Data (see handleRefreshClick /
     * handlePullClick, both driven by flattenAllMasterDataRecords +
     * ExcelService.appendManualBatch). Kept as a self-contained
     * utility in case an isNew-only incremental refresh mode is
     * wanted again later.
     *

     * - First sync for this connection (data.isFirstSync === true,
     *   set by the backend from the connection's last_synced_at):
     *   the sheet has nothing to preserve yet, so this just delegates
     *   to writeMasterData and writes everything, same as before.
     * - Every later refresh: only records the backend flagged isNew
     *   (created after the connection's last successful sync) are
     *   written, appended after exactly one blank row below the
     *   current end of the sheet's data — never inside existing
     *   rows, never overwriting them. isUpdated records are
     *   deliberately left alone: an append-only ledger has no way to
     *   "update" a row it already wrote without touching prior data,
     *   which this must never do.
     * - No new records: nothing is written at all, so nothing here
     *   ever produces a duplicate of a previously-appended record.
     *
     * @param {string} provider - Active ERP provider ("quickbooks" | "xero")
     * @param {object} data - Master Data Pull response, including isFirstSync
     * @returns {Promise<number>} number of records written (new-only on a
     *   later refresh; the full total on a first sync)
     */
    async appendNewMasterData(provider, data) {
        if (!data) return 0;

        if (data.isFirstSync) {
            await ExcelService.writeMasterData(provider, data);
            return countAllMasterDataRecords(data);
        }

        const flatQueue = flattenChangedMasterDataRecords(data);
        if (flatQueue.length === 0) return 0;

        await ExcelService.appendManualBatch(provider, flatQueue);
        return flatQueue.length;
    },

    /**
     * Appends one already-sliced batch of tagged records (see
     * flattenChangedMasterDataRecords / batchDataLoader.js's manual
     * batch queue) below whatever is already on the "1.Master_Data"
     * sheet — never clearing or reordering existing rows, exactly
     * like appendNewMasterData's writer, just parametrized on a
     * pre-sliced list instead of a raw Master Data Pull response.
     *
     * This is what Pull Master Data / Refresh Schedule call directly
     * with a single click's worth of records (one server-paginated
     * page — up to 10 per entity); appendNewMasterData (the un-batched
     * "write everything new" path) now delegates here too, passing
     * its whole flattened list in one call, so both paths share one
     * writer.
     *
     * @param {string} provider - Active ERP provider ("quickbooks" | "xero")
     * @param {{category: string, record: object}[]} batch
     * @returns {Promise<number>} number of records written
     */
    async appendManualBatch(provider, batch) {
        if (!batch || batch.length === 0) return 0;

        const fallbackOrgName = provider === "quickbooks" ? "QuickBooks Company" : "Xero Organisation";
        const orgGroupsMap = new Map();
        const getOrCreateGroup = (name) => {
            const cleanName = (name && name !== "Default" && name !== "Default Organization") ? name : fallbackOrgName;
            const key = cleanName.trim();
            if (!orgGroupsMap.has(key)) {
                orgGroupsMap.set(key, { name: cleanName, accounts: [], classes: [], locations: [], entities: [] });
            }
            return orgGroupsMap.get(key);
        };

        for (const { category, record } of batch) {
            if (!record) continue;

            if (category === "company") {
                // Seeds an org group from the ERP's company/organization
                // list even when it has no accounts/classes/locations/
                // customers/vendors of its own yet, so the org's name
                // row still appears — mirrors writeMasterData's original
                // "Group Companies" step. Only flattenAllMasterDataRecords
                // (the Pull Master Data queue) ever produces this
                // category; the Refresh Schedule queue never does.
                getOrCreateGroup(record.name || fallbackOrgName);
                continue;
            }

            const orgKey = record.clientName || record.clientId || fallbackOrgName;

            if (category === "account") {
                getOrCreateGroup(orgKey).accounts.push(record);
            } else if (category === "class") {
                getOrCreateGroup(orgKey).classes.push(record);
            } else if (category === "location") {
                getOrCreateGroup(orgKey).locations.push(record);
            } else if (category === "customer") {
                getOrCreateGroup(orgKey).entities.push({
                    name: record.name || record.DisplayName || record.Name || "",
                    type: "Customer",
                    id: record.id || record.Id || record.ContactID || "",
                    status: record.active !== undefined ? (record.active ? "Active" : "Inactive") : "Active"
                });
            } else if (category === "vendor") {
                getOrCreateGroup(orgKey).entities.push({
                    name: record.name || record.DisplayName || record.Name || "",
                    type: "Vendor",
                    id: record.id || record.Id || record.ContactID || "",
                    status: record.active !== undefined ? (record.active ? "Active" : "Inactive") : "Active"
                });
            }
        }

        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItem("1.Master_Data");

            // The sheet holds FIVE independent tables side by side, each
            // in its own column block, all sharing the single header row
            // written once by setupWorkbookSheets (row 1 — never
            // rewritten here). Each block therefore has its own next-free
            // row: appending is per block, not per sheet.
            //
            // This matters because the backend drains one API at a time
            // (Accounts fully, then Classes, then Locations, then
            // Customers/Vendors). A batch of 10 classes must land at the
            // top of the Classes block, directly under the last class
            // already written — NOT below the accounts that happen to
            // occupy rows further down the sheet. Positioning off the
            // whole sheet's used range would push each new table down
            // past every earlier one, leaving a staircase of blank rows
            // above it.
            //
            // Each block's next row is derived from that block's own used
            // range, so it stays correct even if the sheet was edited
            // between refreshes, and rows land contiguously with no gap
            // between batches.
            const BLOCKS = {
                company:   { first: "A", last: "B" },
                accounts:  { first: "D", last: "L" },
                classes:   { first: "N", last: "Q" },
                locations: { first: "S", last: "V" },
                entities:  { first: "X", last: "AB" }
            };

            const usedByBlock = {};
            for (const [key, { first, last }] of Object.entries(BLOCKS)) {
                const used = sheet.getRange(`${first}2:${last}10000`).getUsedRangeOrNullObject();
                used.load(["rowIndex", "rowCount", "isNullObject"]);
                usedByBlock[key] = used;
            }

            // Org names already present in column A. The org row is an
            // identity row for the client, not a per-batch banner: it is
            // written once, on whichever batch first introduces that org,
            // and every later batch skips it. Deriving that from the
            // sheet itself (rather than a "is this the first batch?"
            // flag) makes it idempotent — a repeated or replayed batch
            // still cannot duplicate the row.
            const existingOrgColumn = sheet.getRange("A2:A10000");
            existingOrgColumn.load("values");

            // Load existing record IDs for deduplication:
            // Accounts ID is in col L (L1 = QBO Account Id), Classes ID in col P (P1 = QBO Class Id),
            // Locations ID in col U (U1 = QBO Location Id), Entities ID in col AA (AA1 = QBO Entity Id)
            const existingAccountsIdRange = sheet.getRange("L2:L10000");
            const existingClassesIdRange = sheet.getRange("P2:P10000");
            const existingLocationsIdRange = sheet.getRange("U2:U10000");
            const existingEntitiesIdRange = sheet.getRange("AA2:AA10000");

            existingAccountsIdRange.load("values");
            existingClassesIdRange.load("values");
            existingLocationsIdRange.load("values");
            existingEntitiesIdRange.load("values");

            await context.sync();

            const extractIds = (range) => new Set(
                (range.values || [])
                    .map(row => (row && row[0] != null ? String(row[0]).trim() : ""))
                    .filter(Boolean)
            );

            const seenAccountIds = extractIds(existingAccountsIdRange);
            const seenClassIds = extractIds(existingClassesIdRange);
            const seenLocationIds = extractIds(existingLocationsIdRange);
            const seenEntityIds = extractIds(existingEntitiesIdRange);

            // usedRange.rowIndex is 0-based and rowCount is a length, so
            // (rowIndex + rowCount) is the 1-based index of the LAST used
            // row; the next free row is one past it. An empty block
            // starts at row 2, directly beneath the header.
            const nextRowFor = (key) => {
                const used = usedByBlock[key];
                return used.isNullObject ? 2 : used.rowIndex + used.rowCount + 1;
            };

            const rowFor = {};
            for (const key of Object.keys(BLOCKS)) rowFor[key] = nextRowFor(key);

            const seenOrgNames = new Set(
                (existingOrgColumn.values || [])
                    .map(row => (row && row[0] != null ? String(row[0]).trim() : ""))
                    .filter(Boolean)
            );

            let totalWrittenRecords = 0;

            /**
             * Writes one block's rows at that block's own next free row
             * and advances it, so consecutive batches stack without gaps.
             */
            const writeBlock = async (key, values) => {
                if (!values.length) return;
                const { first, last } = BLOCKS[key];
                const startRow = rowFor[key];

                await writeRowsInBatches(context, sheet, first, last, startRow, values, values.map(() => false));

                const written = sheet.getRange(`${first}${startRow}:${last}${startRow + values.length - 1}`);
                written.format.font.size = 11;
                written.format.wrapText = true;

                rowFor[key] = startRow + values.length;
                if (key !== "company") {
                    totalWrittenRecords += values.length;
                }
            };

            for (const [, group] of orgGroupsMap) {
                const orgName = group.name;

                // Header/identity row for this client — only if the sheet
                // doesn't already carry it.
                if (!seenOrgNames.has(orgName.trim())) {
                    await writeBlock("company", [[orgName, orgName]]);
                    seenOrgNames.add(orgName.trim());
                }

                // Filter out accounts already existing in Excel
                const newAccounts = group.accounts.filter(a => {
                    const id = String(a.id || a.Id || a.AccountID || "").trim();
                    if (!id) return true;
                    if (seenAccountIds.has(id)) return false;
                    seenAccountIds.add(id);
                    return true;
                });

                await writeBlock("accounts", newAccounts.map(a => [
                    orgName,
                    a.acctNum || a.code || a.AcctNum || a.Code || "",
                    a.name || a.Name || "",
                    a.accountType || a.type || a.AccountType || a.Type || "",
                    a.accountSubType || a.description || a.AccountSubType || "",
                    a.classification || a.Classification || "",
                    a.fullyQualifiedName || a.name || a.Name || "",
                    a.active !== undefined ? (a.active ? "Active" : "Inactive") : "Active",
                    a.id || a.Id || a.AccountID || ""
                ]));

                // Filter out classes already existing in Excel
                const newClasses = group.classes.filter(c => {
                    const id = String(c.id || c.Id || "").trim();
                    if (!id) return true;
                    if (seenClassIds.has(id)) return false;
                    seenClassIds.add(id);
                    return true;
                });

                await writeBlock("classes", newClasses.map(c => [
                    orgName,
                    c.name || c.Name || "",
                    c.id || c.Id || "",
                    c.active !== undefined ? (c.active ? "Active" : "Inactive") : "Active"
                ]));

                // Filter out locations already existing in Excel
                const newLocations = group.locations.filter(l => {
                    const id = String(l.id || l.Id || "").trim();
                    if (!id) return true;
                    if (seenLocationIds.has(id)) return false;
                    seenLocationIds.add(id);
                    return true;
                });

                await writeBlock("locations", newLocations.map(l => [
                    orgName,
                    l.name || l.Name || "",
                    l.id || l.Id || "",
                    l.active !== undefined ? (l.active ? "Active" : "Inactive") : "Active"
                ]));

                // Filter out entities (Customers/Vendors) already existing in Excel
                const newEntities = group.entities.filter(e => {
                    const id = String(e.id || "").trim();
                    if (!id) return true;
                    if (seenEntityIds.has(id)) return false;
                    seenEntityIds.add(id);
                    return true;
                });

                await writeBlock("entities", newEntities.map(e => [
                    orgName, e.name, e.type, e.id, e.status
                ]));
            }

            sheet.getRange("A:AB").format.columnWidth = 115;

            await context.sync();
            return totalWrittenRecords;
        });

        return batch.length;
    },

    /**
     * Clears only the "1.Master_Data" sheet's data range (A2:AB10000)
     * in place — the same clear writeMasterData does at the start of
     * a full rewrite, pulled out on its own so the manual, batched
     * Pull Master Data flow can clear exactly once at the start of a
     * pull cycle and then append each batch afterward via
     * appendManualBatch, instead of clearing (and losing prior
     * batches) on every click.
     */
    async clearMasterDataRange() {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItem("1.Master_Data");
            const clearRange = sheet.getRange("A2:AB10000");
            clearRange.clear("All");
            await context.sync();
        });
    },

    async clearMasterData() {
        await Excel.run(async (context) => {
            const masterSheet = context.workbook.worksheets.getItemOrNullObject("1.Master_Data");
            const inputSheet = context.workbook.worksheets.getItemOrNullObject("2.Input");
            const sheets = context.workbook.worksheets;
            sheets.load("items/name");
            await context.sync();

            let otherSheetExists = false;
            for (let i = 0; i < sheets.items.length; i++) {
                const name = sheets.items[i].name;
                if (name !== "1.Master_Data" && name !== "2.Input") {
                    otherSheetExists = true;
                    sheets.items[i].activate();
                    break;
                }
            }
            if (!otherSheetExists) {
                context.workbook.worksheets.add("Sheet1").activate();
            }
            if (!masterSheet.isNullObject) masterSheet.delete();
            if (!inputSheet.isNullObject) inputSheet.delete();
            await context.sync();
        });
    },

    async stampLastRefreshed(timestamp) {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItem("1.Master_Data");
            const cell = sheet.getRange("V1");
            cell.values = [[`Last Refreshed: ${timestamp}`]];
            cell.format.font.bold = true;
            cell.format.font.color = "white";
            await context.sync();
        });
    },

    /**
     * Scaffolds required sheets (1.Master_Data and 2.Input) and styles the header rows across 5 colored sections.
     * @param {string} provider - Active ERP provider
     */
    async setupWorkbookSheets(provider) {
        await Excel.run(async (context) => {
            let masterSheet = context.workbook.worksheets.getItemOrNullObject("1.Master_Data");
            let inputSheet = context.workbook.worksheets.getItemOrNullObject("2.Input");
            await context.sync();

            if (masterSheet.isNullObject) masterSheet = context.workbook.worksheets.add("1.Master_Data");
            if (inputSheet.isNullObject) inputSheet = context.workbook.worksheets.add("2.Input");

            masterSheet.activate();
            await context.sync();

            const idLabel = provider === "quickbooks" ? "QBO" : "Xero";

            // Generate spreadsheet header columns from A1 to AB1
            const headerRange = masterSheet.getRange("A1:AB1");
            const headers = [
                [
                    "Client ID", "Client Name", "",
                    "Client ID", "Account Code", "Account Name", "Account Type", "Account Sub-Type", "Classification", "Fully Qualified Name", "Status", `${idLabel} Account Id`, "",
                    "Client ID", "Class Name", `${idLabel} Class Id`, "Status", "",
                    "Client ID", "Location Name", `${idLabel} Location Id`, "Status", "",
                    "Client ID", "Entity Name", "Entity Type", `${idLabel} Entity Id`, "Status"
                ]
            ];

            headerRange.clear();
            headerRange.values = headers;

            // Format Navy Blue Section 1 Header Range (A1:B1)
            const purpleRange1 = masterSheet.getRange("A1:B1");
            purpleRange1.format.fill.color = "#1B224C";
            purpleRange1.format.font.color = "white";
            purpleRange1.format.font.bold = true;
            purpleRange1.format.horizontalAlignment = "Center";
            purpleRange1.format.verticalAlignment = "Center";
            purpleRange1.format.borders.getItem("EdgeBottom").style = Excel.BorderLineStyle.Continuous;

            // Format Navy Blue Section 2 Header Range (D1:L1)
            const blueRange2 = masterSheet.getRange("D1:L1");
            blueRange2.format.fill.color = "#1F4E79";
            blueRange2.format.font.color = "white";
            blueRange2.format.font.bold = true;
            blueRange2.format.horizontalAlignment = "Center";
            blueRange2.format.verticalAlignment = "Center";
            blueRange2.format.borders.getItem("EdgeBottom").style = Excel.BorderLineStyle.Continuous;

            // Format Green Section 3 & 4 Header Ranges (N1:Q1, S1:V1)
            const greenRanges = ["N1:Q1", "S1:V1"];
            for (const range of greenRanges) {
                const r = masterSheet.getRange(range);
                r.format.fill.color = "#0F7546";
                r.format.font.color = "white";
                r.format.font.bold = true;
                r.format.horizontalAlignment = "Center";
                r.format.verticalAlignment = "Center";
                r.format.borders.getItem("EdgeBottom").style = Excel.BorderLineStyle.Continuous;
            }

            // Format Dark Blue Section 5 Header Range (X1:AB1)
            const purpleRange5 = masterSheet.getRange("X1:AB1");
            purpleRange5.format.fill.color = "#1B224C";
            purpleRange5.format.font.color = "white";
            purpleRange5.format.font.bold = true;
            purpleRange5.format.horizontalAlignment = "Center";
            purpleRange5.format.verticalAlignment = "Center";
            purpleRange5.format.borders.getItem("EdgeBottom").style = Excel.BorderLineStyle.Continuous;

            // Configure spacer columns - clear formats
            const spacers = ["C:C", "M:M", "R:R", "W:W"];
            for (const spacer of spacers) {
                const col = masterSheet.getRange(spacer);
                col.clear("Formats");
            }

            headerRange.format.rowHeight = 28;
            headerRange.format.font.size = 11;
            headerRange.format.wrapText = true;

            // Set all columns in the range to column width 115
            masterSheet.getRange("A:AB").format.columnWidth = 115;

            masterSheet.freezePanes.unfreeze();
            masterSheet.getRange("A2").select();

            await context.sync();
        });
    }
};

export { ExcelService };
