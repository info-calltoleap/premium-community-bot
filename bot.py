import discord
from discord.ext import commands
from google.oauth2 import service_account
from googleapiclient.discovery import build
import os
import logging
from dotenv import load_dotenv

# 載入 .env 文件
load_dotenv()

# 讀取環境變數
DISCORD_TOKEN = os.getenv('DISCORD_TOKEN')
if DISCORD_TOKEN is None:
    raise ValueError("DISCORD_TOKEN is not set. Please check your .env file.")

# 設定日誌
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# 設定 Discord bot
intents = discord.Intents.default()
intents.messages = True  # 開啟訊息內容 intent
intents.guilds = True
intents.members = True  # 開啟成員相關 intent

client = commands.Bot(command_prefix='!', intents=intents)

# Google Sheets 設定
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SERVICE_ACCOUNT_FILE = 'premium-community-bot-api2-16d62f6a46ef.json'
spreadsheet_id = '1qEMc17L8-5GIkmuJs9qvJrhXIchg3ytJQhtOJfoknq0'
range_name = 'Sheet1!A2:D'  # 從第二行開始讀取數據

credentials = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES)
service = build('sheets', 'v4', credentials=credentials)

@client.event
async def on_ready():
    logger.info(f'Logged in as {client.user}!')

@client.event
async def on_member_join(member):
    try:
        # 檢查是否有 DM Channel
        dm_channel = member.dm_channel
        if dm_channel is None:
            dm_channel = await member.create_dm()

        # 發送歡迎訊息
        await dm_channel.send(
            "Welcome to our Discord! This is the verification bot for the Premium Members Hub. "
            "If your subscription plan includes this service, please reply with the email address you used to purchase the plan."
        )
        logger.info(f"Sent DM to {member.name} with verification instructions.")

        def check(msg):
            logger.info(f"Received message: {msg.content}")  # Debug output
            return msg.author == member and msg.channel == dm_channel

        # 等待使用者回應
        message = await client.wait_for('message', timeout=120.0, check=check)

        email = message.content.strip()
        logger.info(f"Received email: {email}")  # Debug output

        # 發送確認收到的訊息
        await dm_channel.send(
            "Your message has been received. We are currently verifying it. Please wait a moment."
        )

        # 讀取 Google Sheets 資料
        sheet = service.spreadsheets()
        result = sheet.values().get(spreadsheetId=spreadsheet_id, range=range_name).execute()
        values = result.get('values', [])

        matched_row_index = None
        for i, row in enumerate(values):
            if len(row) > 2 and row[2].strip() == email:  # Column C is Email
                matched_row_index = i
                break

        if matched_row_index is not None:
            matched_row = values[matched_row_index]
            if len(matched_row) > 3 and matched_row[3].strip() == 'used':  # Column D is Status
                await dm_channel.send(
                    "Sorry, this email has already been used. Please double check you entered your email correctly, or contact our support team at info@calltoleap.com"
                )
                logger.info(f"Email {email} has already been used.")
            else:
                role = discord.utils.get(member.guild.roles, name='Premium Member')
                if role:
                    await member.add_roles(role)
                    await dm_channel.send(
                        "Congratulations! Your verification has been successful. Welcome to our Premium Members Hub!"
                    )
                    logger.info(f"Email {email} verified and role added.")

                    # 更新 Google Sheets
                    if len(matched_row) < 4:
                        matched_row.append('used')  # 確保 matched_row 有足夠的長度
                    else:
                        matched_row[3] = 'used'
                    
                    update_range = f'Sheet1!A{matched_row_index + 2}:D{matched_row_index + 2}'
                    body = {
                        'values': [matched_row]
                    }
                    sheet.values().update(spreadsheetId=spreadsheet_id, range=update_range, valueInputOption='RAW', body=body).execute()
                    logger.info(f"Updated email {email} as used in Google Sheets.")
                else:
                    await dm_channel.send("Role 'Premium Member' not found.")
                    logger.warning("Role 'Premium Member' not found.")
        else:
            await dm_channel.send(
                "Sorry, your verification failed. Please double check you entered your email correctly, or contact our support team at info@calltoleap.com"
            )
            logger.warning(f"Email {email} not found in the list.")

    except Exception as e:
        await dm_channel.send("An error occurred while processing your request.")
        logger.error(f"Error occurred: {e}")

client.run(DISCORD_TOKEN)
