import matplotlib.pyplot as plt
import openpyxl
import pandas as pd

df = pd.read_excel("output_fii_stock.xlsx", sheet_name="Equity", engine='openpyxl')

print(df)

plt.figure(figsize=(20, 6))  # Optional: Adjust figure size
plt.plot(df['Date'], df['FII Net Trade'], marker='o', color='b', label='Trade')  # Line plot

# Add labels and title
plt.xlabel('Date')
plt.ylabel('FII Net Trade')
plt.title('Trade Over Time')

# Show grid
plt.grid(True)

# Optional: Add a legend
plt.legend()

# Show the plot
plt.xticks(rotation=45)  # Rotate x-axis labels for better readability
plt.tight_layout()  # Adjust layout to avoid overlap
plt.show()
