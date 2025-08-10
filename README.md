# Introduction
I wrote this program with the help of GPT and use it frequently, but I'm not planning to support any features beyond what's already there.
Feel free to fork this repo and make any changes to it you like!
If you have questions, ask GPT or your AI of choice, you'll get a quicker and more intelligent response than from me ;-)

# You will need to add your google credentials:
Get your credentials.json file that has the credentials to your google account
I won't explain how to get the credentials file, but it's easy enough and takes ~5 minutes with the help of your AI

# Setup
1. copy the example_categories.yaml file and rename it to: categories.yaml
2. edit this file to fit your calendar categories
3. copy the example_blacklist_dates file and rename it to: blacklist_dates
4. edit this file to fit days you want removed from the data
5. Install the necessary python packages in your local environment
6. Run the program

# Extra Info
The very first time you run the script, you will be prompted to authenticate with google. This is necessary to create your personal token.pickle file, which gets used by the google api to query your calendar events.
After a while the token.pickle file will expire, so if your script was working fine just yesterday and you didn't make any changes, just delete the token.pickle file and rerun your script, that will prompt you again to authenticate and your script should be working again.