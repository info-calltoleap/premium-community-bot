import discord
from discord.ext import commands
from google.oauth2 import service_account
from googleapiclient.discovery import build
import os
import logging
from dotenv import load_dotenv
import asyncio
import re

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
intents.message_content = True  # 确保启用了 message_content intent

client = commands.Bot(command_prefix='!', intents=intents)

# Google Sheets 設定
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SERVICE_ACCOUNT_FILE = 'premium-community-bot-api2-16d62f6a46ef.json'
spreadsheet_id = '1qEMc17L8-5GIkmuJs9qvJrhXIchg3ytJQhtOJfoknq0'
range_name = 'Sheet1!A3:E'  # 從第三行開始讀取數據

credentials = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES)
service = build('sheets', 'v4', credentials=credentials)

# 用戶嘗試次數記錄
attempts = {}

@client.event
async def on_ready():
    logger.info(f'Logged in as {client.user}!')
    client.loop.create_task(check_cancellation_emails())

@client.event
async def on_message(message):
    if message.author == client.user:
        return

    channel_id = 1277314657698451506  # 使用您的公共頻道ID

    if message.channel.id != channel_id:
        return

    cleaned_content = message.content.strip().replace('\u200b', '').replace('\n', '').replace('\r', '')
    cleaned_content = re.sub(r'\[.*?\]\(.*?\)', '', cleaned_content)

    email_regex = r'^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$'
    if re.match(email_regex, cleaned_content):
        email = cleaned_content
        member = message.author

        # 限制嘗試次數
        user_id = str(member.id)
        if user_id not in attempts:
            attempts[user_id] = 0

        if attempts[user_id] >= 3:
            await message.channel.send(
                f"{member.mention}, you have reached the maximum number of attempts. Please contact our support team at info@calltoleap.com."
            )
            return

        attempts[user_id] += 1

        sheet = service.spreadsheets()
        result = sheet.values().get(spreadsheetId=spreadsheet_id, range=range_name).execute()
        values = result.get('values', [])

        matched_row_index = None
        for i, row in enumerate(values):
            if len(row) > 2 and row[2].strip().lower() == email.lower():  # Column C is Email
                matched_row_index = i
                break

        if matched_row_index is not None:
            matched_row = values[matched_row_index]

            if len(matched_row) < 5:
                matched_row.extend([''] * (5 - len(matched_row)))

            if len(matched_row) > 3 and matched_row[3].strip() == 'used':  # Column D is Status
                await message.channel.send(
                    f"{member.mention}, sorry, this email has already been used. Please double check you entered your email correctly, or contact our support team at info@calltoleap.com"
                )
            else:
                role = discord.utils.get(member.guild.roles, name='Trade Alerts')
                if role:
                    await member.add_roles(role)
                    await message.channel.send(
                        f"{member.mention}, congratulations! Your verification has been successful. Welcome to our Premium Members Hub!"
                    )

                    matched_row[3] = 'used'  # 標記為已使用
                    matched_row[4] = str(member.id)  # 添加 Discord ID

                    cancellation_range = 'Sheet1!H2:J'
                    cancellation_result = sheet.values().get(spreadsheetId=spreadsheet_id, range=cancellation_range).execute()
                    cancellation_values = cancellation_result.get('values', [])

                    for j, cancel_row in enumerate(cancellation_values):
                        if len(cancel_row) > 2 and cancel_row[2].strip().lower() == email.lower():  # 第十列是 Email
                            cancel_range = f'Sheet1!H{j + 2}:J{j + 2}'
                            clear_body = {
                                'values': [['', '', '']]
                            }
                            sheet.values().update(spreadsheetId=spreadsheet_id, range=cancel_range, valueInputOption='RAW', body=clear_body).execute()

                    update_range = f'Sheet1!A{matched_row_index + 3}:E{matched_row_index + 3}'  # 修正行數
                    body = {
                        'values': [matched_row]
                    }
                    sheet.values().update(spreadsheetId=spreadsheet_id, range=update_range, valueInputOption='RAW', body=body).execute()
                else:
                    await message.channel.send(f"{member.mention}, role 'Trade Alerts' not found.")
        else:
            await message.channel.send(
                f"{member.mention}, sorry, your verification failed. Please double check you entered your email correctly, or contact our support team at info@calltoleap.com"
            )
    else:
        await message.channel.send(
            f"{message.author.mention}, please enter a valid email address for verification."
        )

# 設置一個任務來定期檢查取消訂閱的電子郵件
async def check_cancellation_emails():
    while True:
        try:
            sheet = service.spreadsheets()

            result = sheet.values().get(spreadsheetId=spreadsheet_id, range=range_name).execute()
            values = result.get('values', [])

            cancellation_range = 'Sheet1!H2:J'
            cancellation_result = sheet.values().get(spreadsheetId=spreadsheet_id, range=cancellation_range).execute()
            cancellation_values = cancellation_result.get('values', [])

            if cancellation_values:
                for j, cancel_row in enumerate(cancellation_values):
                    if len(cancel_row) > 2:
                        cancel_email = cancel_row[2].strip()

                        email_matched_index = None
                        for i, row in enumerate(values):
                            if len(row) > 2 and row[2].strip().lower() == cancel_email.lower():
                                email_matched_index = i
                                break

                        if email_matched_index is not None:
                            matched_row = values[email_matched_index]

                            if len(matched_row) < 5:
                                matched_row.extend([''] * (5 - len(matched_row)))

                            cancel_range = f'Sheet1!H{j + 2}:J{j + 2}'
                            clear_body = {
                                'values': [['', '', '']]
                            }
                            sheet.values().update(spreadsheetId=spreadsheet_id, range=cancel_range, valueInputOption='RAW', body=clear_body).execute()

                            matched_row = ['', '', '', '', '']

                            if len(row) > 4 and row[4].strip():
                                discord_id = int(row[4].strip())
                                guild = client.get_guild(768962332524937258)
                                member = guild.get_member(discord_id)

                                if member:
                                    role = discord.utils.get(guild.roles, name='Trade Alerts')
                                    if role:
                                        await member.remove_roles(role)

                            update_range = f'Sheet1!A{email_matched_index + 3}:E{email_matched_index + 3}'
                            body = {
                                'values': [matched_row]
                            }
                            sheet.values().update(spreadsheetId=spreadsheet_id, range=update_range, valueInputOption='RAW', body=body).execute()

            await asyncio.sleep(60)
        except Exception as e:
            logger.error(f"Error checking cancellation emails: {e}")

client.run(DISCORD_TOKEN)
