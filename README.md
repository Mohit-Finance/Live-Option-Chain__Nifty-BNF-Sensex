# 📊 Live Option Chain Dashboard (Python + Upstox API + Excel)

A **real-time option chain streaming system** built using **Python** and **Upstox Broker API**, with **tick-by-tick updates directly into Excel**.

This project streams live option chain data for **NIFTY, BANKNIFTY, and SENSEX**, calculates key option metrics (Greeks, IV, OI, depth, PCR), and displays everything in an Excel workbook that updates live — making it ideal for **analysis, strategy testing, and automation research**.

> ⚠️ This is a **data & analytics project**, not a trading recommendation system.

---

## 🚀 Key Features

### 🔴 Live Tick-by-Tick Updates
- Real-time option chain streaming using **Upstox WebSocket API**
- Updates Excel **without refresh or reload**
- Near zero latency (broker-dependent)

### 📈 Supported Indices
- **NIFTY 50**
- **BANKNIFTY**
- **SENSEX**

Each index has its **own Excel sheet** with synchronized updates.

---

## 🧠 Data Captured (Per Strike)

### 📌 Core Option Data
- Option Price (LTP)
- Previous Close
- Volume
- Open Interest (OI)
- Change in OI
- Bid / Ask Price
- Market Depth

### 📐 Option Greeks (Live)
- Delta
- Gamma
- Vega
- Theta
- Implied Volatility (IV)

### 📊 Derived Metrics
- OI Change %
- Price Change %
- Put-Call Ratio (PCR)
- Depth-based Strength
- Strike-wise positioning

---

## 🧾 Excel Dashboard Structure

### 📂 Sheets
- `nifty` – NIFTY option chain
- `bnf` – BANKNIFTY option chain
- `sensex` – SENSEX option chain

---

### 🧩 Column Description (Example: NIFTY Sheet)

| Column | Description |
|------|------------|
| expiry | Option expiry date |
| prev_close_CE | Previous close (Call) |
| delta_CE | Delta (Call) |
| gamma_CE | Gamma (Call) |
| vega_CE | Vega (Call) |
| theta_CE | Theta (Call) |
| iv_CE | Implied Volatility (Call) |
| volume_CE | Traded volume (Call) |
| OI_CE | Open Interest (Call) |
| chg_OI_CE | Change in OI (Call) |
| strike | Strike price |
| ltp_CE | Last traded price (Call) |
| bid_CE | Best bid (Call) |
| ask_CE | Best ask (Call) |
| depth_CE | Market depth (Call) |
| ltp_PE | Last traded price (Put) |
| bid_PE | Best bid (Put) |
| ask_PE | Best ask (Put) |
| volume_PE | Traded volume (Put) |
| OI_PE | Open Interest (Put) |
| chg_OI_PE | Change in OI (Put) |
| iv_PE | Implied Volatility (Put) |
| delta_PE | Delta (Put) |
| gamma_PE | Gamma (Put) |
| theta_PE | Theta (Put) |
| strike_PCR | Strike-wise PCR |
| nifty_spot | Spot price |
| India_VIX | India VIX |

📌 **Color coding in Excel**
- 🟢 Green → Positive / bullish change
- 🔴 Red → Negative / bearish change

---

## 🖼 Sample Output

### Live NIFTY Option Chain (Excel)
![NIFTY Option Chain](./assets/nifty_option_chain.png)

*(BANKNIFTY and SENSEX sheets follow the same structure)*

---

## 🛠 Tech Stack

| Layer | Technology |
|-----|-----------|
| Language | Python 3.10+ |
| Broker API | Upstox API (REST + WebSocket) |
| Excel Integration | xlwings |
| Data Handling | pandas, numpy |
| WebSocket | websockets, asyncio |
| Auth | pyotp (2FA automation) |
| Visualization | Excel conditional formatting |

---

## 📦 Python Libraries Used

```text
pandas
numpy
xlwings
websockets
asyncio
upstox_client
protobuf
pyotp
matplotlib
openpyxl
