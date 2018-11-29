# FortiOS Virtual IP

Recently I was involved in a project that required many VIP creations with specific ports.

I couldn't find any API or Ansible Module to do that, so I decided to create this script to automate this process.


Note* I use .xlsx as the input file, due to the fact that the project that I was involved used this format to store their documents !!
You also will notice that I set the default protocol on TCP which vividly is changeable by your needs!



How does it work?
It will ask your credentials first, then reads your documents that contains the source/destination of IP and Ports that you want to publish(look at MyDoc.xlsx as an example).

Then it connects to your devices(FortiOS) via SSH and runs VIP commands to create VIPs and groups.
