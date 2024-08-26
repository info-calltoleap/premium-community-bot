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
intents.messages = True  # Enable message content intent
intents.guilds = True
intents.members = True  # Enable member-related intent
intents.message_content = True  # Ensure message_content intent is enabled

client = commands.Bot(command_prefix='!', intents=intents)

# Google Sheets setup
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SERVICE_ACCOUNT_FILE = 'premium-community-bot-api2-16d62f6a46ef.json'
spreadsheet_id = '1qEMc17L8-5GIkmuJs9qvJrhXIchg3ytJQhtOJfoknq0'
range_name = 'Sheet1!A3:E'  # Start reading data from the third row

credentials = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES)
service = build('sheets', 'v4', credentials=credentials)

# User attempt tracking
attempts = {}

@client.event
async def on_ready():
    logger.info(f'Logged in as {client.user}!')
    client.loop.create_task(check_cancellation_emails())

@client.event
async def on_message(message):
    if message.author == client.user:
        return

    if message.content.startswith('!reset_attempts') or message.content.startswith('!reset_role_attempts'):
        await client.process_commands(message)
        return

    channel_id = 1277314657698451506  # Your public channel ID

    if message.channel.id != channel_id:
        return

    cleaned_content = message.content.strip().replace('\u200b', '').replace('\n', '').replace('\r', '')
    cleaned_content = re.sub(r'\[.*?\]\(.*?\)', '', cleaned_content)

    email_regex = r'^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$'
    if re.match(email_regex, cleaned_content):
        email = cleaned_content
        member = message.author

        # Limit attempts
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

                    matched_row[3] = 'used'  # Mark as used
                    matched_row[4] = str(member.id)  # Add Discord ID

                    cancellation_range = 'Sheet1!H2:J'
                    cancellation_result = sheet.values().get(spreadsheetId=spreadsheet_id, range=cancellation_range).execute()
                    cancellation_values = cancellation_result.get('values', [])

                    for j, cancel_row in enumerate(cancellation_values):
                        if len(cancel_row) > 2 and cancel_row[2].strip().lower() == email.lower():  # Column J is Email
                            cancel_range = f'Sheet1!H{j + 2}:J{j + 2}'
                            clear_body = {
                                'values': [['', '', '']]
                            }
                            sheet.values().update(spreadsheetId=spreadsheet_id, range=cancel_range, valueInputOption='RAW', body=clear_body).execute()

                    update_range = f'Sheet1!A{matched_row_index + 3}:E{matched_row_index + 3}'  # Adjust row number
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

# Set up a task to periodically check for cancellation emails
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
                                        if role in member.roles:
                                            await member.remove_roles(role)
                                            logger.info(f"Removed 'Trade Alerts' role from {member.name}.")
                                        else:
                                            logger.warning(f"User {member.name} does not have the 'Trade Alerts' role.")

                                    await member.send(
                                        "Your subscription has been canceled. If you have any questions, please contact our support team at info@calltoleap.com"
                                    )
                                else:
                                    logger.warning(f"User with ID {discord_id} not found in the server.")

                            update_range = f'Sheet1!A{email_matched_index + 3}:E{email_matched_index + 3}'
                            body = {
                                'values': [matched_row]
                            }
                            sheet.values().update(spreadsheetId=spreadsheet_id, range=update_range, valueInputOption='RAW', body=body).execute()

            await asyncio.sleep(60)
        except Exception as e:
            logger.error(f"Error checking cancellation emails: {e}")

# Reset the attempts of a specific user
@client.command()
@commands.has_permissions(administrator=True)
async def reset_attempts(ctx, member: discord.Member):
    user_id = str(member.id)
    if user_id in attempts:
        attempts[user_id] = 0
        await ctx.send(f"{member.mention}, please re-verify your email.")
    else:
        await ctx.send(f"{member.mention} has not made any verification attempts.")

# Reset the attempts of all members in a specific role
@client.command()
@commands.has_permissions(administrator=True)
async def reset_role_attempts(ctx, role_name: str):
    # Get the role from the server
    role = discord.utils.get(ctx.guild.roles, name=role_name)
    if not role:
        await ctx.send(f"Role `{role_name}` does not exist.")
        return

    # Loop through all members with the role and reset their attempts
    reset_count = 0
    for member in role.members:
        user_id = str(member.id)
        if user_id in attempts:
            attempts[user_id] = 0
            reset_count += 1
            await ctx.send(f"{member.mention}, please re-verify your email.")

    await ctx.send(f"Reset verification attempts for {reset_count} members with the role `{role_name}`.")

@reset_attempts.error
@reset_role_attempts.error
async def reset_error(ctx, error):
    if isinstance(error, commands.MissingPermissions):
        await ctx.send(f"{ctx.author.mention}, you do not have permission to perform this action.")

client.run(DISCORD_TOKEN)
