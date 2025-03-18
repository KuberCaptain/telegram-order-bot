# Telegram Order Bot Documentation

## Overview
The **Telegram Order Bot** is a Python-based Telegram bot that facilitates order placement, product listing, and blacklist management. It interacts with Excel files to store product details, order history, and banned users.

## Features
- **/start & /help** – Displays available commands.
- **/Magaz** – Lists available products.
- **/Order** – Allows users to place an order.
- **/blacklist** – Shows the blacklist of banned users.
- **Admin Commands**:
  - **/orders** – View order history (admin only).

## Installation
### Prerequisites
Ensure you have Python 3.8+ installed. You also need to install dependencies:
```bash
pip install -r requirements.txt
```

### Setting Up the Bot
1. Clone the repository:
```bash
git clone https://github.com/KuberCaptain/telegram-order-bot.git
cd telegram-order-bot
```
2. Create a virtual environment and activate it:
```bash
python -m venv venv
source venv/bin/activate  # Linux/macOS
venv\Scripts\activate    # Windows
```
3. Install required packages:
```bash
pip install -r requirements.txt
```
4. Obtain a Telegram bot token from [@BotFather](https://t.me/BotFather) and replace it in `main.py`:
```python
TOKEN = "YOUR_TELEGRAM_BOT_TOKEN"
```

## Usage
### Running the Bot
To start the bot, execute:
```bash
python main.py
```

### Available Commands
| Command        | Description                        |
|---------------|------------------------------------|
| /start, /help | Displays available commands       |
| /Magaz        | Shows available products         |
| /Order        | Initiates an order process       |
| /blacklist    | Lists banned users              |
| /orders       | (Admin only) View all orders     |

## Order Process
1. User sends **/Order**.
2. Bot asks for the product name.
3. User inputs the product name.
4. Bot asks for the quantity.
5. Order is processed and stored in `orders.xlsx`.
6. An admin receives a notification about the new order.

## File Structure
- **`main.py`** – Main bot logic.
- **`requirements.txt`** – Dependencies.
- **`products.xlsx`** – Stores product information.
- **`orders.xlsx`** – Stores order history.
- **`blacklist.xlsx`** – Stores banned users.

## Admin Functionality
Admins (specified by user ID in `main.py`) can:
- View order history using `/orders`.
- Receive real-time notifications for new orders.

## Deployment
For a production-ready setup:
- Run the bot inside a `systemd` service or a `tmux`/`screen` session.
- Use a hosting provider with a stable connection.
- Consider integrating a database instead of Excel files for scalability.

## Troubleshooting
### Bot doesn’t start
- Ensure the correct **TOKEN** is set in `main.py`.
- Check if dependencies are installed (`pip install -r requirements.txt`).

### Command not recognized
- Ensure the bot has restarted after any modifications.

### Database not updating
- Ensure `orders.xlsx`, `products.xlsx`, and `blacklist.xlsx` exist in the bot directory.

## License
This project is licensed under the MIT License.

## Author
Developed by [KuberCaptain](https://github.com/KuberCaptain).
