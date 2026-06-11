# Bank Performance Dashboard

This dashboard was built to give a bank a single, at-a-glance view of its **deposits, lending, transaction flows, and customer base**across branches. It turns raw account- and transaction-level records into the metrics a branch manager or analyst actually needs, total deposits, loan exposure, account activity, and the all-important **Loan-to-Deposit Ratio**. The final dashboard lives in [`bank_dashboard.xlsx`](bank_dashboard.xlsx).

The following Excel skills were used to build it:

- **Charts:** Column, bar, line, and pie charts driven directly off summary tables.
- **Formulas & Functions:** `SUMIFS()`, `COUNTIFS()`, `COUNTA()`, and `IF()` for all KPIs and aggregations.
- **Dashboard Design:** KPI cards, a consistent navy / pink / blue theme, and a clean grid layout.
- **Dynamic Linking:** A dedicated `Calc` sheet feeds every chart, so the whole dashboard recalculates when source data changes.

## Questions the Dashboard Answers

1. How much does the bank hold in deposits vs. loans?
2. Which branches hold the most balance, and how are accounts distributed?
3. How do deposits and withdrawals trend month over month?
4. How is the customer base split across Retail, SME, and Corporate segments?
5. What is each branch's **Loan-to-Deposit Ratio (LDR)**, a key liquidity indicator?

## The Dataset

The workbook contains sample data across three working sheets plus a calculation layer:

| Sheet | Contents |
|-------|----------|
| `Accounts` | 320 accounts, customer, branch, segment, account type, open date, balance, status |
| `Transactions` | 12 months of deposits / withdrawals per branch |
| `Calc` | Summary tables (aggregations) that feed every chart |
| `Dashboard` | The final visual layer, KPI cards + 5 charts |

> Deposits are stored as **positive**balances and loans as **negative**balances, which lets a single `Balance` column drive both deposit and loan metrics.

---

## Building the Dashboard

### KPI Cards

The four header cards summarize the whole bank in one row.

```excel
Total Deposits   =SUMIFS(Accounts!G:G, Accounts!G:G, ">0")
Total Loans      =-SUMIFS(Accounts!G:G, Accounts!E:E, "Loan")
Total Accounts   =COUNTA(Accounts!A:A)-1
Active Accounts  =COUNTIFS(Accounts!H:H, "Active")
```

- **Formula Purpose:** `SUMIFS` with a `">0"` criterion isolates deposits; the leading `-` on the loans formula flips the negative loan balances to a positive figure.
- **Design Choice:** Large white numbers on solid color blocks so the headline metrics read instantly.

### Balance by Branch

```excel
=SUMIFS(Accounts!$G:$G, Accounts!$C:$C, A2)
```

- **Formula Purpose:** Sums every account balance belonging to a given branch.
- **Excel Feature:** Column chart for quick ranking of branches by total balance.

### Balance by Account Type

```excel
=SUMIFS(Accounts!$G:$G, Accounts!$E:$E, E2)
```

- **Formula Purpose:** Totals balances per account type (Savings, Current, Fixed Deposit, Loan).
- **Design Choice:** Pie chart to show how the balance mix is composed as a share of the whole.

### Monthly Deposits vs Withdrawals

```excel
Deposits     =SUMIFS(Transactions!$C:$C, Transactions!$A:$A, K2)
Withdrawals  =SUMIFS(Transactions!$D:$D, Transactions!$A:$A, K2)
```

- **Formula Purpose:** Rolls the per-branch transaction rows up to a single bank-wide figure for each month.
- **Excel Feature:** Two-series line chart to compare inflow vs outflow trends over 12 months.

### Accounts by Segment

```excel
=COUNTIFS(Accounts!$D:$D, H2)
```

- **Formula Purpose:** Counts how many accounts fall into each customer segment (Retail / SME / Corporate).
- **Excel Feature:** Horizontal bar chart for an easy visual comparison of segment sizes.

---

## Featured Calculation: Loan-to-Deposit Ratio (LDR)

The **LDR**measures how much of a branch's deposit base is lent out — a core liquidity and risk indicator. A lower ratio means more conservative lending relative to deposits.

**Per branch**(in the `Calc` sheet):

```excel
Deposits  =SUMIFS(Accounts!$G:$G, Accounts!$C:$C, O2, Accounts!$G:$G, ">0")
Loans     =-SUMIFS(Accounts!$G:$G, Accounts!$C:$C, O2, Accounts!$E:$E, "Loan")
LDR       =IF(P2=0, 0, Q2/P2)
```

**Overall**(Dashboard KPI box):

```excel
=-SUMIFS(Accounts!G:G, Accounts!E:E, "Loan") / SUMIFS(Accounts!G:G, Accounts!G:G, ">0")
```

- **Formula Purpose:** Loans divided by deposits, per branch and bank-wide. The branch version uses **two criteria**in one `SUMIFS`,  branch name *and* sign/type, to isolate exactly the right balances.
- **Error Handling:** The `IF(P2=0, 0, …)` guard prevents a `#DIV/0!` error if a branch has no deposits.
- **Visual:** Plotted as a percentage-formatted column chart so branches can be compared against each other (and, optionally, a target threshold).

---

## Conclusion

This project demonstrates how a handful of `SUMIFS` / `COUNTIFS` aggregations on a clean `Calc` layer can power a fully dynamic banking dashboard. Because every chart and KPI references formulas — not hardcoded values, swapping the sample rows in `Accounts` and `Transactions` for real data refreshes the entire dashboard automatically.

### To Use With Real Data
1. Replace the rows in the `Accounts` and `Transactions` sheets, keeping the same column structure.
2. Recalculate (the `Calc` tables and all charts update on their own).
3. Review the KPI cards and LDR chart for the refreshed period.
