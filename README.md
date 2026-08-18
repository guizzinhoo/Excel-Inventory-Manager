# 📦 StockPilot

A desktop inventory management system developed in Python for small businesses that manage their products through Microsoft Excel.

Instead of manually editing spreadsheets every day, this application provides a complete interface to manage products, sales and reports while keeping Excel as the data source.

---

## Features

- 📦 Product registration
- 🔍 Product search
- 🛒 Sales registration
- 🔄 Stock replenishment
- 📊 Sales reports
- 📈 Revenue analysis
- 📋 Sales history
- ⚠ Low stock alerts
- 🖥 Desktop graphical interface
- 📄 Excel integration

---

## Technologies

- Python
- Pandas
- Openpyxl
- CustomTkinter
- Matplotlib

---

## Project Structure

```
StockPilot/
│
├── Backend.py
├── interface.py
├── AvonNatura.xlsx
├── requirements.txt
└── README.md
```

---

## How it Works

The application uses an Excel workbook as its database.

Every operation performed through the interface automatically updates the spreadsheet.

The system can:

- register new products
- update stock quantities
- register sales
- calculate available inventory
- calculate revenue
- generate sales history
- identify products with low stock

---

## Main Screens

- Inventory
- Register Sale
- Add Product
- Restock
- Reports
- Sales History
- Low Stock

---

## Reports

The application generates management information such as:

- Total revenue
- Total items sold
- Top selling products
- Revenue distribution
- Inventory status

---

## Motivation

This project was created to simplify inventory management for a small family business that previously relied entirely on manual Excel editing.

The objective was to automate repetitive tasks while keeping the workflow familiar for the user.

---

## Future Improvements

- SQLite database support
- User authentication
- PDF report export
- Dashboard with KPIs
- Barcode reader support

---

## License

MIT
