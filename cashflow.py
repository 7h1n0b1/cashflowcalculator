import pandas as pd
from datetime import datetime
import matplotlib.pyplot as plt

# initialize variables
cash_flow = {}
current_month = None
income = 0
expenses = 0
expense_by_narration = {}

# read Excel file
df = pd.read_excel('transactions.xlsx')

# iterate over each row
for index, row in df.iterrows():
    # parse date and transaction amounts
    date_str = row['Date']
    narration = row['Narration']
    withdrawal = row['Withdrawal Amt.']
    deposit = row['Deposit Amt.']
    amount = 0
    if pd.notna(withdrawal):
        amount -= float(withdrawal)
    if pd.notna(deposit):
        amount += float(deposit)

    # parse date using the correct format string
    date = datetime.strptime(date_str, '%d/%m/%y')

    # calculate total income and expenses for each month
    if current_month != date.strftime('%Y-%m'):
        if current_month:
            cash_flow[current_month] = income - expenses
        current_month = date.strftime('%Y-%m')
        income = 0
        expenses = 0

    # add transaction to income or expenses based on amount
    if amount > 0:
        income += amount
    else:
        expenses -= amount
        # add expense to narration category
        if narration in expense_by_narration:
            expense_by_narration[narration] += amount
        else:
            expense_by_narration[narration] = amount

# add final month's cash flow
cash_flow[current_month] = income - expenses

# create DataFrame of cash flow
cash_flow_df = pd.DataFrame(cash_flow.items(), columns=['Month', 'Cash Flow'])

# write DataFrame to Excel sheet
with pd.ExcelWriter('cash_flow.xlsx') as writer:
    cash_flow_df.to_excel(writer, index=False)

    # create pie chart of expenses by narration
    if expense_by_narration:
        # remove negative values
        expense_by_narration = {k: v for k, v in expense_by_narration.items() if v >= 0}
        if expense_by_narration:
            narration_df = pd.DataFrame(expense_by_narration.items(), columns=['Narration', 'Expense'])
            narration_df.plot(kind='pie', y='Expense', labels=narration_df['Narration'], autopct='%1.1f%%', legend=False)
            plt.title('Expenses by Narration')
            plt.ylabel('')
            plt.tight_layout()
            plt.savefig('expenses_by_narration.png')
            plt.close()
            
            # write pie chart to Excel sheet
            workbook = writer.book
            worksheet = workbook.add_worksheet('Expenses by Narration')
            expenses_chart = workbook.add_chart({'type': 'pie'})
            expenses_chart.add_series({
                'name': 'Expenses',
                'categories': ['Expenses by Narration', 1, 0, len(narration_df), 0],
                'values': ['Expenses by Narration', 1, 1, len(narration_df), 1],
            })
            expenses_chart.set_title({'name': 'Expenses by Narration'})
            expenses_chart.set_legend({'position': 'bottom'})
            worksheet.insert_chart('A1', expenses_chart)

