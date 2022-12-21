# PyPostman4RPG

At January 2020 I took a part in Role Play Game. Before the game, master decided to allow letter exchange between the players.
But, the players in real world mustn't know who plays any role in the game.
This program sent approximately 800 letters, between 30 players, during one month!

Masters was satisfied!

Masters and I imagine the next workflow:
1. Players fill the simple form for each letter in the google forms with several fields:
    - sender name (the name of your role)
    - recipient of the letter
    - letter file, it can be TXT or PDF

2. Masters open a google sheet, read each letter and solve, who must be a recipient. For simplify the process, i added
additional sheet with addressbook. It allows to make a drop-down list with recipients. Note-bene: remember about master arbitrariness!
If player define recipient, (if you game allow it) the letter can be stolen, sent to another player, or be lost

3. Masters click on the special link, which execute the postman program or wait. Everyday, Crontab starts the program at specified time.

4. The postman-program connects to google-spreadsheets throw google API:
   4.1 check new post-rows which was addressed by masters
   4.2 fetch email from list with addressbook
   4.3 fetch letter file from google-drive.
   4.5 if letter in PDF, clear 'Author' in metadata
   4.6 Send letter to recipient
   4.7 fill the send time in the table

## Before use

You must recive API key from google and specify different params in the script.