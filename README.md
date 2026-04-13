**📌 AfriTrade Mini ERP System**

AfriTrade is a lightweight Enterprise Resource Planning (ERP) system built using Microsoft Excel VBA. It is designed to manage business transactions, automate currency conversion to Nigerian Naira (NGN), and provide real-time dashboard insights.

**🚀 Features**
📊 Sales & Purchases Management
💱 Real-time Currency Conversion (USD, EUR, GBP → NGN)
🧾 Automated Transaction Recording via UserForms
📈 Dynamic Dashboard for KPI Tracking
⚙️ Modular VBA Architecture (modFX, modSales, modDashboard)
🔄 API Integration for Live Exchange Rates

**🛠️ Technologies Used**
Microsoft Excel
VBA (Visual Basic for Applications)
REST API (Exchange Rate API)
JSON Parsing (VBA-JSON)

**🧩 System Architecture**

The system is structured into modules:

modFX → Handles currency conversion and API integration
modsalesform → Handles user input
modDashboard → Updates KPIs and reports
modSales→ Sales Order Processing
modPurchases→ Stock Replenishment

**⚙️ How It Works**
User opens the system and clicks a button to launch the input form
Transaction details are entered
System fetches live exchange rates via API
Amount is converted to NGN automatically
Data is stored in the worksheet
Errors are logged in the errorlog worksheet
Dashboard updates instantly with new totals
