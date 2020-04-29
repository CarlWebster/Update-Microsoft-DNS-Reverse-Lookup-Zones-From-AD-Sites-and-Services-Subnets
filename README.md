# UpdateReverseZonesFromSubnets
Update AD DNS Reverse Lookup Zones from AD Sites &amp; Services Subnets

	Script to get subnets from Sites & Services, see if a matching AD DNS Reverse Lookup 
	Zone exists, and if not, create the reverse zone.

	This script was created for a client who granted permission to share with the 
	community.

	This script reads the list of Subnets in AD Sites & Services, checks if a matching DNS 
	Reverse Lookup Zone exists, if the zone does not exist, attempts to create it.
	
	The Reverse Zones created are created with a Replication Scope of "Forest" and Dynamic 
	Update set to "Secure". By default, Aging and Scavenging is enabled with the default 
	7 days No-refresh and Refresh intervals.

	The script requires at least PowerShell version 4 but runs best in version 5.

	This script requires Domain Admin rights and an elevated PowerShell session.

	Creates an output file named UpdateReverseZonesFromSubnetsScriptResults_YYYYMMDDHHSS.txt.

	You do NOT have to run this script on a domain controller. This script was developed 
	and run from a Windows 10 VM.

	To run the script from a workstation, RSAT is required.
