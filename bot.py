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

@client.event
async def on_ready():
    logger.info(f'Logged in as {client.user}!')
    client.loop.create_task(check_cancellation_emails())

@client.event
async def on_message(message):
    # 確保不會回應機器人自己的消息
    if message.author == client.user:
        return

    # 設定公共頻道ID
    channel_id = 1277314657698451506  # 使用您的公共頻道ID

    if message.channel.id != channel_id:
        return

    # 打印收到的原始消息内容（用于调试）
    logger.info(f"Raw message content: {repr(message.content)} from {message.author} in channel {message.channel}")

    # 清理消息内容，移除不可见字符、换行符等
    cleaned_content = message.content.strip().replace('\u200b', '').replace('\n', '').replace('\r', '')
    cleaned_content = re.sub(r'\[.*?\]\(.*?\)', '', cleaned_content)  # 去除Markdown链接格式
    logger.info(f"Cleaned message content: '{cleaned_content}'")

    # 判斷是否為有效的電子郵件
    email_regex = r'^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$'
    if re.match(email_regex, cleaned_content):
        email = cleaned_content
        member = message.author

        await message.channel.send(
            f"{member.mention}, your message has been received. We are currently verifying it. Please wait a moment."
        )

        # 讀取 Google Sheets 資料
        sheet = service.spreadsheets()
        result = sheet.values().get(spreadsheetId=spreadsheet_id, range=range_name).execute()
        values = result.get('values', [])

        logger.info(f"Google Sheets data: {values}")

        matched_row_index = None
        for i, row in enumerate(values):
            if len(row) > 2 and row[2].strip().lower() == email.lower():  # Column C is Email
                matched_row_index = i
                break

        if matched_row_index is not None:
            matched_row = values[matched_row_index]
            
            # 確保 matched_row 的長度足夠
            if len(matched_row) < 5:
                matched_row.extend([''] * (5 - len(matched_row)))

            if len(matched_row) > 3 and matched_row[3].strip() == 'used':  # Column D is Status
                await message.channel.send(
                    f"{member.mention}, sorry, this email has already been used. Please double check you entered your email correctly, or contact our support team at info@calltoleap.com"
                )
                logger.info(f"Email {email} has already been used.")
            else:
                role = discord.utils.get(member.guild.roles, name='Trade Alerts')
                if role:
                    await member.add_roles(role)
                    await message.channel.send(
                        f"{member.mention}, congratulations! Your verification has been successful. Welcome to our Premium Members Hub!"
                    )
                    logger.info(f"Email {email} verified and role added.")

                    # 更新 Google Sheets
                    matched_row[3] = 'used'  # 標記為已使用
                    matched_row[4] = str(member.id)  # 添加 Discord ID

                    # 獲取取消訂閱的資料
                    cancellation_range = 'Sheet1!H2:J'
                    cancellation_result = sheet.values().get(spreadsheetId=spreadsheet_id, range=cancellation_range).execute()
                    cancellation_values = cancellation_result.get('values', [])

                    # 檢查第十列的 email 是否匹配
                    for j, cancel_row in enumerate(cancellation_values):
                        if len(cancel_row) > 2 and cancel_row[2].strip().lower() == email.lower():  # 第十列是 Email
                            cancel_range = f'Sheet1!H{j + 2}:J{j + 2}'
                            clear_body = {
                                'values': [['', '', '']]  # 清空第八到第十列的資料
                            }
                            sheet.values().update(spreadsheetId=spreadsheet_id, range=cancel_range, valueInputOption='RAW', body=clear_body).execute()
                            logger.info(f"Cleared cancellation data for email {email} in columns H-J.")

                    # 更新 Google Sheets
                    update_range = f'Sheet1!A{matched_row_index + 2}:E{matched_row_index + 2}'
                    body = {
                        'values': [matched_row]
                    }
                    sheet.values().update(spreadsheetId=spreadsheet_id, range=update_range, valueInputOption='RAW', body=body).execute()
                    logger.info(f"Updated email {email} as used and added Discord ID {member.id} in Google Sheets.")
                else:
                    await message.channel.send(f"{member.mention}, role 'Trade Alerts' not found.")
                    logger.warning("Role 'Trade Alerts' not found.")
        else:
            await message.channel.send(
                f"{member.mention}, sorry, your verification failed. Please double check you entered your email correctly, or contact our support team at info@calltoleap.com"
            )
            logger.warning(f"Email {email} not found in the list.")
    else:
        logger.warning(f"Invalid email format received: '{cleaned_content}'")
        await message.channel.send(
            f"{message.author.mention}, please enter a valid email address for verification."
        )

# 設置一個任務來定期檢查取消訂閱的電子郵件
async def check_cancellation_emails():
    while True:
        try:
            # 重新获取 Google Sheets 服务实例
            sheet = service.spreadsheets()

            # 获取所有数据行
            result = sheet.values().get(spreadsheetId=spreadsheet_id, range=range_name).execute()
            values = result.get('values', [])

            # 读取消订订阅的电子邮件
            cancellation_range = 'Sheet1!H2:J'
            cancellation_result = sheet.values().get(spreadsheetId=spreadsheet_id, range=cancellation_range).execute()
            cancellation_values = cancellation_result.get('values', [])

            if cancellation_values:
                for j, cancel_row in enumerate(cancellation_values):
                    if len(cancel_row) > 2:  # 确保有Email US栏位
                        cancel_email = cancel_row[2].strip()

                        # 在主要列表中找到匹配的电子邮件
                        email_matched_index = None
                        for i, row in enumerate(values):
                            if len(row) > 2 and row[2].strip().lower() == cancel_email.lower():  # Column C 是 Email
                                email_matched_index = i
                                break

                        if email_matched_index is None:
                            matched_row = values[email_matched_index]

                            # 確保 matched_row 的長度足夠以清空数据
                            if len(matched_row) < 5:
                                matched_row.extend([''] * (5 - len(matched_row)))

                            # 删除第十列的email匹配的行的第八列到第十列的数据
                            cancel_range = f'Sheet1!H{j + 2}:J{j + 2}'
                            clear_body = {
                                'values': [['', '', '']]  # 清空第八到第十列的資料
                            }
                            sheet.values().update(spreadsheetId=spreadsheet_id, range=cancel_range, valueInputOption='RAW', body=clear_body).execute()
                            logger.info(f"Cleared data for cancellation email {cancel_email} in columns H-J.")

                            # 清空匹配到的行的第一列到第五列的数据
                            matched_row = ['', '', '', '', '']  # 清空該行的前五列

                            # 获取 Discord ID 并移除角色
                            if len(row) > 4 and row[4].strip():
                                discord_id = int(row[4].strip())
                                guild = client.get_guild(768962332524937258)  # 使用你的服务器 ID
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
                            logger.info(f"Cleared all data for {cancel_email} in columns A-E.")

            await asyncio.sleep(60)  # 每分钟检查一次
        except Exception as e:
            logger.error(f"Error checking cancellation emails: {e}")

client.run(DISCORD_TOKEN)
