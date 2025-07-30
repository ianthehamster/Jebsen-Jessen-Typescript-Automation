# Excel Trial Balance Consolidation Script

This script automates the **consolidation of multiple trial balance (TB) worksheets** in an Excel workbook using Office Scripts (TypeScript). It handles **currency conversion, account code normalization, tax expense adjustments, and Orbitax entity mapping**.

---

## Features

- **Excludes helper sheets** (e.g., `Steps`, `General`, `Orbitax Entity Codes`, etc.) automatically.
- **Currency conversion**:
  - Reads conversion rates from `MAS and Wise Exchange Rates` sheet.
  - Converts all amounts into **SGD** using matching currency codes.
- **Trial balance consolidation**:
  - Aggregates all TB data from multiple sheets starting at row 9 (`A9:H`).
  - Handles missing or invalid account codes (generates random 6‑digit fallback codes starting with `1`).
- **Account sign normalization**:
  - Flips signs for revenue/expense accounts (`4`, `5`, `6`, `7`) except specific tax accounts.
- **Entity mapping with Orbitax**:
  - Maps entity codes/names from the `General` sheet to Orbitax codes/names.
- **Generates two output sheets**:
  1. `TB with Positive Tax Exp` – all amounts in SGD with tax expenses positive.
  2. `TB with Negative Tax Exp` – tax expense accounts flipped to negative for GloBE adjustments.

---

## Input Sheet Requirements

- **Currency rates sheet (`MAS and Wise Exchange Rates`)**
  - Currency codes in `B10:L10`
  - SGD rates in `B11:L11`

- **General mapping sheet (`General`)**
  - Orbitax entity code in column `B`
  - Orbitax entity name in column `D`
  - Jebsen entity short name in column `G`

- **Trial balance sheets**
  - Data starts from row 9
  - Must follow structure:
    ```
    A: Entity code
    B: Account code
    C: Account name
    D: Currency
    E: Amount
    (Columns F–H ignored)
    ```

---

## How It Works

1. **Collect all sheets** excluding helper and TB sheets.
2. **Extract currency codes/rates** and determine exchange rate for each TB sheet’s currency.
3. **Transform each row**:
   - Remove numeric prefixes from account names.
   - Flip signs for non‑tax expense accounts starting with `4–7`.
   - Convert amounts to SGD.
   - Generate random 6‑digit fallback code for missing account codes.
4. **Aggregate all rows** into `TB with Positive Tax Exp`.
5. **Map Orbitax codes/names** from `General` sheet onto consolidated TB.
6. **Duplicate and modify** the positive TB to create `TB with Negative Tax Exp`:
   - Flip specified tax expense accounts (`72000`, `73000`, `73005`, `73010`, `73020`) to negative amounts.

---

## Usage

1. Open the workbook in Excel on the web.
2. Add this script in **Automate > Code Editor**.
3. Ensure required sheets (`MAS and Wise Exchange Rates`, `General`, TB sheets) are present.
4. Run the script via **Automate > Run**.

---

## Configuration

- **Excluded Sheets**: Update `excludeSheetNames` array if new helper sheets are added.
- **Tax Accounts**: Modify `taxExpenseAccounts` and `taxExpenseAccountsToBeFlippedToNegative` arrays for new tax codes.
- **Currency Codes**: Ensure `MAS and Wise Exchange Rates` layout matches expectations (`B10:L11`).

---

## Output

- **TB with Positive Tax Exp**
- **TB with Negative Tax Exp**

Both sheets contain 5 columns:
Entity Code | Entity Name | Account Code | Account Name | Amount (SGD)

---

## Error Handling

- Random codes are generated for accounts without valid codes.
- Existing output sheets are deleted before creating new ones.
- Sheet name conflicts are avoided by deletion prior to creation.

