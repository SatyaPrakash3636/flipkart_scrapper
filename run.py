from flipkart.flipkart import Flipkart

with Flipkart() as bot:
    bot.landing_page()
    bot.login_page()
    bot.select_seller()
    bot.earn_more()
