Server Assistant is a tool which allows easy and automated 
administration of a Half-Life Server.

Originally intended to be released publically for a small fee, 
I never had the time to polish it off and fix the various bugs.

So, I am releasing the complete source code. This project is 
over 4 years old, so some of the code may seem strange, and
it probably wont all flow together vary well.

The most complex thing in this project, i'd say, is the AvScript
scripting language. I did this one for fun, but it's pretty 
good... the complete thing I wrote, including the parser and 
the math functions and so on. Kudos to anyone who can make out
the infinite recursive functions and actually get it to do 
something better!

SA talks to HLDS by using the UDP RCON protocol, and HLDS talks to
SA by sending it the log files via UDP using the logaddress command.

SAC connects directly to SA via TCP, so you'll need to open a port on
your firewall. You can configure SA by editing assistant.cfg.

SA can also be an RCON Proxy for your admins.

The default user/pass is:

user: admin
pass: changeme

Enjoy! And remember, if you get any use out of this, please donate to

avpaypal@cyberwyre.com

Thanks!

- Avatar-X
avcode@cyberwyre.com

Note: This is distributed under the GNU General Public License.

