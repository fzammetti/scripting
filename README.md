# scripting
A collection of VBS scripts, batch files and shell scripts 

Being largely a Windows guy, I often find myself writing some VBS script or batch file to do things. You'll find some of those here. I Linux/Unix too though so you may find some shell scripts here too.

* **archive_subdirectories.vbs** - 7z's all subdirectories in the current directory, password-protecting them
  
* **dns_ips.vbs** - Gets an IP address for a number of domains and writes them out to a file.  This is helpful if your ISP's DNS server goes down, this way you have the IPs to use (and if you schedule this to run once a night like I do then you're usually good to go when needed).  Of course, nowadays, most of us have Google, Cloudflare and THEN (maybe) ISP DNS servers configured, so it's less of a concern, but I suppose there's still some use for this.
  
* **extract_archives_to_subdirectories.vbs** - extract all .7z archives in the current directory to subdirectories (there was a time when 7-Zip couldn't do this by default I believe, which is why I wrote this)
  
* **list_installed_software.vbs** - list all installed software and saves it to a file
