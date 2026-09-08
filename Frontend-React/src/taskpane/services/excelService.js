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
        if (typeof Excel === "undefined") {
            throw new Error("Excel API is not available in this environment.");
        }
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItem("1.Master_Data");
            const usedRange = sheet.getUsedRangeOrNullObject();
            usedRange.load(["rowCount", "isNullObject"]);
            await context.sync();

            if (!usedRange.isNullObject && usedRange.rowCount > 1) {
                const clearRowCount = Math.max(usedRange.rowCount, 500);
                sheet.getRange(`A2:AB${clearRowCount + 50}`).clear("All");
            } else {
                sheet.getRange("A2:AB100").clear("All");
            }

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

            const rawCompanies = Array.isArray(data.company)
                ? data.company
                : (data.company ? [data.company] : []);

            for (const c of rawCompanies) {
                if (c) getOrCreateGroup(c.name || c.id, c.name);
            }

            if (data.accounts) {
                for (const a of data.accounts) {
                    const nameKey = a.clientName || a.clientId || fallbackOrgName;
                    const group = getOrCreateGroup(a.clientId, nameKey);
                    group.accounts.push(a);
                }
            }

            if (data.classes) {
                for (const c of data.classes) {
                    const nameKey = c.clientName || c.clientId || fallbackOrgName;
                    const group = getOrCreateGroup(c.clientId, nameKey);
                    group.classes.push(c);
                }
            }

            if (data.locations) {
                for (const l of data.locations) {
                    const nameKey = l.clientName || l.clientId || fallbackOrgName;
                    const group = getOrCreateGroup(l.clientId, nameKey);
                    group.locations.push(l);
                }
            }

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

            let currentRow = 2;

            for (const [key, group] of orgGroupsMap) {
                const orgName = group.name;
                const accCount = group.accounts.length;
                const classCount = group.classes.length;
                const locCount = group.locations.length;
                const entityCount = group.entities.length;
                const maxRows = Math.max(1, accCount, classCount, locCount, entityCount);

                sheet.getRange(`A${currentRow}:B${currentRow}`).values = [[orgName, orgName]];

                if (accCount > 0) {
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

                currentRow += maxRows + 2;
            }

            const formatEndRow = Math.max(currentRow, 50);
            const dataRange = sheet.getRange(`A2:AB${formatEndRow}`);
            dataRange.format.font.size = 11;
            dataRange.format.wrapText = true;
            sheet.getRange("A:AB").format.columnWidth = 115;

            await context.sync();
        });
    },

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

    async appendManualBatch(provider, batch) {
        if (!batch || batch.length === 0) return 0;
        if (typeof Excel === "undefined") {
            throw new Error("Excel API is not available in this environment.");
        }

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
            const BLOCKS = {
                company:   { first: "A", last: "B" },
                accounts:  { first: "D", last: "L" },
                classes:   { first: "N", last: "Q" },
                locations: { first: "S", last: "V" },
                entities:  { first: "X", last: "AB" }
            };

            const overallUsed = sheet.getUsedRangeOrNullObject();
            overallUsed.load(["rowCount", "isNullObject"]);
            await context.sync();
            const scanLimit = (!overallUsed.isNullObject && overallUsed.rowCount > 1) ? Math.max(overallUsed.rowCount + 500, 2000) : 2000;

            const usedByBlock = {};
            for (const [key, { first, last }] of Object.entries(BLOCKS)) {
                const used = sheet.getRange(`${first}2:${last}${scanLimit}`).getUsedRangeOrNullObject();
                used.load(["rowIndex", "rowCount", "isNullObject"]);
                usedByBlock[key] = used;
            }

            const existingOrgColumn = sheet.getRange(`A2:A${scanLimit}`);
            existingOrgColumn.load("values");

            const existingAccountsIdRange = sheet.getRange(`L2:L${scanLimit}`);
            const existingClassesIdRange = sheet.getRange(`P2:P${scanLimit}`);
            const existingLocationsIdRange = sheet.getRange(`U2:U${scanLimit}`);
            const existingEntitiesIdRange = sheet.getRange(`AA2:AA${scanLimit}`);

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

                if (!seenOrgNames.has(orgName.trim())) {
                    await writeBlock("company", [[orgName, orgName]]);
                    seenOrgNames.add(orgName.trim());
                }

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

    async clearMasterDataRange() {
        if (typeof Excel === "undefined") return;
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItemOrNullObject("1.Master_Data");
            sheet.load("isNullObject");
            await context.sync();
            if (sheet.isNullObject) return;

            const used = sheet.getUsedRangeOrNullObject();
            used.load(["rowCount", "isNullObject"]);
            await context.sync();

            if (!used.isNullObject && used.rowCount > 1) {
                const clearRowCount = Math.max(used.rowCount, 500);
                sheet.getRange(`A2:AB${clearRowCount + 50}`).clear("All");
            } else {
                sheet.getRange("A2:AB100").clear("All");
            }
            await context.sync();
        });
    },

    async clearMasterData() {
        if (typeof Excel === "undefined") return;
        await Excel.run(async (context) => {
            const masterSheet = context.workbook.worksheets.getItemOrNullObject("1.Master_Data");
            const inputSheet = context.workbook.worksheets.getItemOrNullObject("2.Input");
            const sheets = context.workbook.worksheets;
            sheets.load("items/name");
            await context.sync();

            if (!masterSheet.isNullObject) {
                try { masterSheet.getRange().clear("All"); } catch (_) {}
            }
            if (!inputSheet.isNullObject) {
                try { inputSheet.getRange().clear("All"); } catch (_) {}
            }
            await context.sync();

            const existingNames = new Set(sheets.items.map(s => s.name));
            let fallbackSheet = sheets.items.find(s => s.name !== "1.Master_Data" && s.name !== "2.Input");

            if (!fallbackSheet) {
                let newName = "Sheet1";
                let counter = 1;
                while (existingNames.has(newName)) {
                    newName = `Sheet${++counter}`;
                }
                fallbackSheet = context.workbook.worksheets.add(newName);
            }

            fallbackSheet.activate();
            await context.sync();

            if (!masterSheet.isNullObject) {
                try { masterSheet.delete(); } catch (_) {}
            }
            if (!inputSheet.isNullObject) {
                try { inputSheet.delete(); } catch (_) {}
            }
            await context.sync();
        });
    },

    async stampLastRefreshed(timestamp) {
        if (typeof Excel === "undefined") return;
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
        if (typeof Excel === "undefined") {
            throw new Error("Excel API is not available in this environment.");
        }
        await Excel.run(async (context) => {
            let masterSheet = context.workbook.worksheets.getItemOrNullObject("1.Master_Data");
            let inputSheet = context.workbook.worksheets.getItemOrNullObject("2.Input");
            await context.sync();

            if (masterSheet.isNullObject) masterSheet = context.workbook.worksheets.add("1.Master_Data");
            if (inputSheet.isNullObject) inputSheet = context.workbook.worksheets.add("2.Input");

            masterSheet.activate();
            await context.sync();

            const idLabel = (provider || "").toLowerCase() === "quickbooks" ? "QBO" : "Xero";

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
            masterSheet.getRange("A:AB").format.columnWidth = 115;

            masterSheet.freezePanes.unfreeze();
            masterSheet.getRange("A2").select();

            await context.sync();
        });
    }
};

export { ExcelService };
