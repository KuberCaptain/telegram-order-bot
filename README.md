# Telegram Order Bot 📦

This bot allows you to manage a store via Telegram:
- 🛒 **View products** (`/Magaz`)
- 📝 **Place orders** (`/Order`)
- ⚠️ **Check the blacklist** (`/ЧС`)

## 🚀 Installation and Setup

### **1. Clone the repository**
```bash
git clone https://github.com/KuberCaptain/telegram-order-bot.git
cd telegram-order-bot
```

### **2. Install dependencies**
```bash
pip install -r requirements.txt
```

### **3. Run the bot**
```bash
python main.py
```

## 🔧 Compile to `.exe`
To create a standalone executable file, run:
```bash
pyinstaller --onefile main.py
```
The compiled file will be located in the `dist/` folder.

## 📂 Project Files
| File            | Description |
|----------------|------------|
| `main.py`      | Main bot script |
| `blacklist.xlsx` | Blacklist of users |
| `products.xlsx` | List of store products |
| `orders.xlsx`   | Order logs |

## 📬 Bot Commands
| Command   | Description |
|-----------|------------|
| `/start`  | Start the bot |
| `/help`   | List available commands |
| `/Magaz`  | Show product catalog |
| `/Order`  | Place an order |
| `/ЧС`     | Check blacklist |

## 🛠 Development
### Run in VS Code:
1. Open the terminal (`Ctrl + ~`)
2. Run:
   ```bash
   python main.py
   ```

## 📝 License
This project is licensed under the **MIT License**.
