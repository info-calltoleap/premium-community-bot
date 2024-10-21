import discord
from discord.ext import commands
from google.oauth2 import service_account
from googleapiclient.discovery import build
import os
import logging
from dotenv import load_dotenv
import asyncio
import re

# Load .env file
load_dotenv()

# Read environment variables
DISCORD_TOKEN = os.getenv('DISCORD_TOKEN')
if DISCORD_TOKEN is None:
    raise ValueError("DISCORD_TOKEN is not set. Please check your .env file.")

# Set up logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Set up Discord bot
intents = discord.Intents.default()
intents.messages = True
intents.guilds = True
intents.members = True
intents.message_content = True

client = commands.Bot(command_prefix='!', intents=intents)

# Google Sheets setup
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SERVICE_ACCOUNT_FILE = 'premium-community-bot-api2-16d62f6a46ef.json'
spreadsheet_id = '1qEMc17L8-5GIkmuJs9qvJrhXIchg3ytJQhtOJfoknq0'
range_name = 'Sheet1!A3:E'  # Start reading data from the third row

credentials = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES)
service = build('sheets', 'v4', credentials=credentials)

@client.event
async def on_ready():
    logger.info(f'Logged in as {client.user}!')
    client.loop.create_task(check_cancellation_emails())

@client.event
async def on_message(message):
    if message.author == client.user:
        return

    channel_id = 1277310796522848266  # Your public channel ID

    # Check if the message is in the verification channel
    if message.channel.id != channel_id:
        return

    # 定義機器人的 Discord 使用者 ID
    bot_ids = [159985870458322944, 1281627943428161536, 1275728804567978005]

    # Skip messages from the specific bot IDs
    if message.author.id in bot_ids:
        logger.info(f"Skipped deleting message from bot {message.author.id}")
        return

    # Remove the message for privacy reasons
    logger.info(f"Deleting message from {message.author.id} in channel {message.channel.id}")
    await message.delete()

    # Validate the email format
    cleaned_content = message.content.strip().replace('\u200b', '').replace('\n', '').replace('\r', '')
    email_regex = r'^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$'
    
    if not re.match(email_regex, cleaned_content):
        await message.channel.send(f"{message.author.mention}, please double-check if you’ve entered your email incorrectly.")
        return

    email = cleaned_content
    member = message.author
    guild = member.guild

    # Check if email already exists in L and M columns (Discord ID and email logging)
    sheet = service.spreadsheets()
    check_range = 'Sheet1!L:M'
    result = sheet.values().get(spreadsheetId=spreadsheet_id, range=check_range).execute()
    existing_values = result.get('values', [])

    email_already_logged = False
    for row in existing_values:
        if len(row) > 1 and row[1].strip().lower() == email.lower():
            email_already_logged = True
            break

    if not email_already_logged:
        # Log Discord ID and email in L and M columns
        new_row = [[str(member.id), email]]
        append_range = 'Sheet1!L:M'
        sheet.values().append(spreadsheetId=spreadsheet_id, range=append_range, valueInputOption='RAW', body={'values': new_row}).execute()
        logger.info(f"Logged {member.name}'s Discord ID and email.")

    # Now check if the email exists in column C for verification
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

        if matched_row[3].strip() == 'used':  # Status is "used"
            await message.channel.send(f"{member.mention}, sorry, this email has already been used.")
        else:
            # Add "General" and "Trade Alerts" roles
            general_role = discord.utils.get(guild.roles, name='General')
            trade_alerts_role = discord.utils.get(guild.roles, name='Trade Alerts')
            await member.add_roles(general_role, trade_alerts_role)

            await message.channel.send(f"{member.mention}, Welcome to our community! Since your monthly plan includes the Member Hub, you now have access to visit these channels!")

            # Update Google Sheet with Discord ID and status "used"
            matched_row[3] = 'used'
            matched_row[4] = str(member.id)

            update_range = f'Sheet1!A{matched_row_index + 3}:E{matched_row_index + 3}'  # Adjust row number
            body = {'values': [matched_row]}
            sheet.values().update(spreadsheetId=spreadsheet_id, range=update_range, valueInputOption='RAW', body=body).execute()

    else:
        # If email is not found in column C, just add "General" role
        general_role = discord.utils.get(guild.roles, name='General')
        await member.add_roles(general_role)
        await message.channel.send(f"{member.mention}, Welcome to our community! Don’t be shy to interact with us, and feel free to ask any questions or join the conversations!")

# Check for cancellation emails every minute
async def check_cancellation_emails():
    while True:
        try:
            sheet = service.spreadsheets()
            result = sheet.values().get(spreadsheetId=spreadsheet_id, range=range_name).execute()
            values = result.get('values', [])

            # Check for cancellation emails in column J
            cancellation_range = 'Sheet1!J3:J'
            cancellation_result = sheet.values().get(spreadsheetId=spreadsheet_id, range=cancellation_range).execute()
            cancellation_values = cancellation_result.get('values', [])

            if cancellation_values:
                for j, cancel_row in enumerate(cancellation_values):
                    if len(cancel_row) > 0:
                        cancel_email = cancel_row[0].strip()

                        # Find and remove corresponding email in column C
                        for i, row in enumerate(values):
                            if len(row) > 2 and row[2].strip().lower() == cancel_email.lower():
                                # Delete corresponding A-E row
                                update_range_a_e = f'Sheet1!A{i + 3}:E{i + 3}'
                                body_a_e = {'values': [['', '', '', '', '']]}
                                sheet.values().update(spreadsheetId=spreadsheet_id, range=update_range_a_e, valueInputOption='RAW', body=body_a_e).execute()

                                # Remove corresponding H-J row
                                update_range_h_j = f'Sheet1!H{j + 3}:J{j + 3}'
                                body_h_j = {'values': [['', '', '']]}  # Clear H, I, J columns
                                sheet.values().update(spreadsheetId=spreadsheet_id, range=update_range_h_j, valueInputOption='RAW', body=body_h_j).execute()

                                # Remove roles from Discord user
                                if len(row) > 4 and row[4].strip():
                                    discord_id = int(row[4].strip())
                                    guild = client.get_guild(768962332524937258)
                                    member = guild.get_member(discord_id)

                                    if member:
                                        trade_alerts_role = discord.utils.get(guild.roles, name='Trade Alerts')
                                        premium_role = discord.utils.get(guild.roles, name='Premium Member')

                                        roles_to_remove = [role for role in [trade_alerts_role, premium_role] if role in member.roles]

                                        if roles_to_remove:
                                            await member.remove_roles(*roles_to_remove)
                                            logger.info(f"Removed roles {[role.name for role in roles_to_remove]} from {member.name}.")
            
            await asyncio.sleep(21600)  # Check every 6 hour
        except Exception as e:
            logger.error(f"Error checking cancellation emails: {e}")

client.run(DISCORD_TOKEN)
