import discord
from discord.ext import commands
from google.oauth2 import service_account
from googleapiclient.discovery import build
import os
import logging
from dotenv import load_dotenv
import asyncio

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
range_name = 'Sheet1!A2:E'  # 從第二行開始讀取數據

credentials = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES)
service = build('sheets', 'v4', credentials=credentials)

@client.event
async def on_ready():
    logger.info(f'Logged in as {client.user}!')
    client.loop.create_task(check_cancellation_emails())

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
                role = discord.utils.get(member.guild.roles, name='Trade Alerts')
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

                    if len(matched_row) < 5:
                        matched_row.append(str(member.id))  # 添加 Discord ID 到第五列
                    else:
                        matched_row[4] = str(member.id)  # 更新 Discord ID

                    update_range = f'Sheet1!A{matched_row_index + 2}:E{matched_row_index + 2}'
                    body = {
                        'values': [matched_row]
                    }
                    sheet.values().update(spreadsheetId=spreadsheet_id, range=update_range, valueInputOption='RAW', body=body).execute()
                    logger.info(f"Updated email {email} as used and added Discord ID {member.id} in Google Sheets.")
                else:
                    await dm_channel.send("Role 'Trade Alerts' not found.")
                    logger.warning("Role 'Trade Alerts' not found.")
        else:
            await dm_channel.send(
                "Sorry, your verification failed. Please double check you entered your email correctly, or contact our support team at info@calltoleap.com"
            )
            logger.warning(f"Email {email} not found in the list.")

    except Exception as e:
        await dm_channel.send("An error occurred while processing your request.")
        logger.error(f"Error occurred: {e}")

# 設置一個任務來定期檢查取消訂閱的電子郵件
async def check_cancellation_emails():
    while True:
        try:
            # 讀取 Google Sheets 中的取消訂閱電子郵件
            cancellation_range = 'Sheet1!H2:J'
            cancellation_result = sheet.values().get(spreadsheetId=spreadsheet_id, range=cancellation_range).execute()
            cancellation_values = cancellation_result.get('values', [])

            if cancellation_values:
                for cancel_row in cancellation_values:
                    if len(cancel_row) > 2:  # 確保有Email US欄位
                        cancel_email = cancel_row[2].strip()

                        # 在主要列表中找到匹配的電子郵件
                        email_matched_index = None
                        for i, row in enumerate(values):
                            if len(row) > 2 and row[2].strip() == cancel_email:  # Column C 是 Email
                                email_matched_index = i
                                break

                        if email_matched_index is not None:
                            matched_row = values[email_matched_index]

                            # 移除"used"狀態
                            if len(matched_row) > 3 and matched_row[3].strip() == 'used':
                                matched_row[3] = ''  # 清除狀態

                            # 取得 Discord ID 並移除角色
                            if len(matched_row) > 4 and matched_row[4].strip():
                                discord_id = int(matched_row[4].strip())
                                guild = client.get_guild(768962332524937258)  # 使用你的伺服器 ID
                                member = guild.get_member(discord_id)

                                if member:
                                    role = discord.utils.get(guild.roles, name='Trade Alerts')
                                    if role:
                                        await member.remove_roles(role)
                                        logger.info(f"Removed 'Trade Alerts' role from {member.name}.")

                            # 更新 Google Sheets
                            update_range = f'Sheet1!A{email_matched_index + 2}:E{email_matched_index + 2}'
                            body = {
                                'values': [matched_row]
                            }
                            sheet.values().update(spreadsheetId=spreadsheet_id, range=update_range, valueInputOption='RAW', body=body).execute()
                            logger.info(f"Cleared 'used' status and updated Discord ID for {cancel_email}.")

            await asyncio.sleep(21600)  # 每6小時檢查一次
        except Exception as e:
            logger.error(f"Error checking cancellation emails: {e}")

client.run(DISCORD_TOKEN)
