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

DISCORD_TOKEN = os.getenv('DISCORD_TOKEN')
if DISCORD_TOKEN is None:
    raise ValueError("DISCORD_TOKEN is not set. Please check your .env file.")

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

intents = discord.Intents.default()
intents.messages = True
intents.guilds = True
intents.members = True
intents.message_content = True

client = commands.Bot(command_prefix='!', intents=intents)

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SERVICE_ACCOUNT_FILE = 'premium-community-bot-api2-16d62f6a46ef.json'
spreadsheet_id = '1qEMc17L8-5GIkmuJs9qvJrhXIchg3ytJQhtOJfoknq0'
range_name = 'Discord!A3:E'

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

    channel_id = 1277310796522848266
    if message.channel.id != channel_id:
        return

    bot_ids = [159985870458322944, 1281627943428161536, 1275728804567978005]
    specific_message = (
        "Ah, noble seeker of knowledge, your gratitude resonates like the sweet sound of a lute in a grand hall! "
        "However, let us not linger too long in pleasantries, for the quest for wisdom, much like the pursuit of the perfect taco, "
        "requires swift action and unwavering determination. If thou hast further inquiries or matters of importance to discuss, "
        "I stand ready, like a steadfast knight clad in shining armor, prepared to assist thee! What dost thou wish to know?"
    )

    if message.author.id == 1281627943428161536 and message.content.strip() == specific_message:
        logger.info(f"Deleting specific message from {message.author.id}")
        await message.delete()
        return

    if message.author.id in bot_ids:
        logger.info(f"Skipped deleting message from bot {message.author.id}")
        return

    logger.info(f"Deleting message from {message.author.id} in channel {message.channel.id}")
    await message.delete()

    cleaned_content = message.content.strip().replace('\u200b', '').replace('\n', '').replace('\r', '')
    email_regex = r'^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$'

    if not re.match(email_regex, cleaned_content):
        await message.channel.send(f"{message.author.mention}, please double-check if you’ve entered your email incorrectly.")
        return

    email = cleaned_content
    member = message.author
    guild = member.guild

    sheet = service.spreadsheets()
    check_range = 'Discord!L:M'
    result = sheet.values().get(spreadsheetId=spreadsheet_id, range=check_range).execute()
    existing_values = result.get('values', [])

    email_already_logged = any(len(row) > 1 and row[1].strip().lower() == email.lower() for row in existing_values)

    if not email_already_logged:
        new_row = [[str(member.id), email]]
        append_range = 'Discord!L:M'
        sheet.values().append(spreadsheetId=spreadsheet_id, range=append_range, valueInputOption='RAW', body={'values': new_row}).execute()
        logger.info(f"Logged {member.name}'s Discord ID and email.")

    result = sheet.values().get(spreadsheetId=spreadsheet_id, range=range_name).execute()
    values = result.get('values', [])

    matched_row_index = next((i for i, row in enumerate(values) if len(row) > 2 and row[2].strip().lower() == email.lower()), None)

    if matched_row_index is not None:
        matched_row = values[matched_row_index]
        while len(matched_row) < 5:
            matched_row.append('')

        if matched_row[3].strip() == 'used':
            await message.channel.send(f"{member.mention}, sorry, this email has already been used.")
        else:
            general_role = discord.utils.get(guild.roles, name='General')
            trade_alerts_role = discord.utils.get(guild.roles, name='Trade Alerts')

            if general_role:
                await member.add_roles(general_role)
            if trade_alerts_role:
                await member.add_roles(trade_alerts_role)

            await message.channel.send(f"{member.mention}, Welcome to our community! Since your plan includes the Member Hub, you now have access to visit these channels!")

            matched_row[3] = 'used'
            matched_row[4] = str(member.id)

            update_range = f'Discord!A{matched_row_index + 3}:E{matched_row_index + 3}'
            body = {'values': [matched_row]}
            sheet.values().update(spreadsheetId=spreadsheet_id, range=update_range, valueInputOption='RAW', body=body).execute()
    else:
        general_role = discord.utils.get(guild.roles, name='General')
        if general_role:
            await member.add_roles(general_role)

        await message.channel.send(f"{member.mention}, Welcome to our community! Don’t be shy to interact with us, and feel free to ask any questions or join the conversations!")

async def check_cancellation_emails():
    while True:
        try:
            sheet = service.spreadsheets()
            result = sheet.values().get(spreadsheetId=spreadsheet_id, range=range_name).execute()
            values = result.get('values', [])

            cancellation_result = sheet.values().get(spreadsheetId=spreadsheet_id, range='Discord!J3:J').execute()
            cancellation_values = cancellation_result.get('values', [])

            if cancellation_values:
                for j, cancel_row in enumerate(cancellation_values):
                    if len(cancel_row) > 0:
                        cancel_email = cancel_row[0].strip()

                        for i, row in enumerate(values):
                            if len(row) > 2 and row[2].strip().lower() == cancel_email.lower():
                                update_range_a_e = f'Discord!A{i + 3}:E{i + 3}'
                                sheet.values().update(spreadsheetId=spreadsheet_id, range=update_range_a_e, valueInputOption='RAW', body={'values': [['', '', '', '', '']]}).execute()

                                update_range_h_j = f'Discord!H{j + 3}:J{j + 3}'
                                sheet.values().update(spreadsheetId=spreadsheet_id, range=update_range_h_j, valueInputOption='RAW', body={'values': [['', '', '']]}).execute()

                                if len(row) > 4 and row[4].strip():
                                    discord_id = int(row[4].strip())
                                    guild = client.get_guild(768962332524937258)
                                    member = guild.get_member(discord_id)

                                    if member:
                                        trade_alerts_role = discord.utils.get(guild.roles, name='Trade Alerts')
                                        premium_role = discord.utils.get(guild.roles, name='Premium Member')
                                        roles_to_remove = [r for r in [trade_alerts_role, premium_role] if r in member.roles]

                                        if roles_to_remove:
                                            await member.remove_roles(*roles_to_remove)
                                            logger.info(f"Removed roles {[r.name for r in roles_to_remove]} from {member.name}.")
            await asyncio.sleep(60)
        except Exception as e:
            logger.error(f"Error checking cancellation emails: {e}")

client.run(DISCORD_TOKEN)
