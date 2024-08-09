import discord
from discord.ext import commands
import os
from dotenv import load_dotenv
import asyncio

# 載入 .env 文件
load_dotenv()

# 讀取環境變數
DISCORD_TOKEN = os.getenv('DISCORD_TOKEN')

if DISCORD_TOKEN is None:
    raise ValueError("DISCORD_TOKEN is not set. Please check your .env file.")

intents = discord.Intents.default()
intents.message_content = True
intents.members = True  # Enable the member intent

bot = commands.Bot(command_prefix="!", intents=intents)

ACCESS_CODE = "calltoleap"  # Set the code you want users to enter

@bot.event
async def on_ready():
    print(f'We have logged in as {bot.user}')

@bot.event
async def on_member_join(member):
    await member.send("Welcome! Please reply with the access code.")
    print(f"Sent access code request to {member}")

    def check(m):
        return m.author == member and m.content == ACCESS_CODE

    try:
        msg = await bot.wait_for('message', check=check, timeout=300)
        print(f"Received code from {member}: {msg.content}")

        role = discord.utils.get(member.guild.roles, name="PrivateChannelAccess")
        if role:
            # Check if the bot has permission to manage roles
            if not member.guild.me.guild_permissions.manage_roles:
                raise PermissionError("Bot does not have 'Manage Roles' permission.")

            # Check if the bot's role is above the role it is trying to add
            if role.position < member.guild.me.top_role.position:
                await member.add_roles(role)
                await member.send("Access granted! You now have access to the private channels.")
                print(f"Granted role to {member}")
            else:
                await member.send("Bot's role is not high enough to assign this role. Please contact support.")
                print(f"Bot's role is too low to grant 'PrivateChannelAccess' to {member}")
        else:
            await member.send("Role not found. Please contact support.")
            print(f"Role 'PrivateChannelAccess' not found for {member}")

    except asyncio.TimeoutError:
        await member.send("Sorry, the code is incorrect. Please try again or contact support.")
        print(f"Timeout occurred for {member}")

    except PermissionError as pe:
        await member.send(str(pe))
        print(f"Permission error: {pe}")

    except Exception as e:
        print(f"An error occurred: {e}")
        await member.send("There was an error processing your request. Please try again later.")

@bot.command()
async def remove_access(ctx, member: discord.Member):
    role = discord.utils.get(ctx.guild.roles, name="PrivateChannelAccess")
    if role:
        # Check if the bot has permission to manage roles
        if not ctx.guild.me.guild_permissions.manage_roles:
            raise PermissionError("Bot does not have 'Manage Roles' permission.")

        # Check if the bot's role is above the role it is trying to remove
        if role.position < ctx.guild.me.top_role.position:
            await member.remove_roles(role)
            await ctx.send(f"Access removed for {member.name}.")
            print(f"Removed role from {member}")
        else:
            await ctx.send("Bot's role is not high enough to remove this role. Please contact support.")
            print(f"Bot's role is too low to remove 'PrivateChannelAccess' from {member}")
    else:
        await ctx.send("Role 'PrivateChannelAccess' not found. Please contact support.")
        print(f"Role 'PrivateChannelAccess' not found for {ctx.author}")

bot.run(DISCORD_TOKEN)
