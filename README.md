# Create EOP Firewall Address Objects

## Description

This Powershell script will pull the current list of Exchange Online Protection (EOP) subnets from Microsoft and create a text file with the CLI commands required to create the appropriate objects and groups for use in firewall policies.

The script currently supports:

- FortiGate
- Palo Alto Networks
- Cisco ASA


## Use Case

Create firewall address objects to be used in EOP lock down firewall policies.

## Execution Dependencies

- Powershell

## Contributor(s)

- Stephen Santos

# Important Execution Notes

## Production Impact

- None

## Resource Impact

- None

## Other Considerations

- None

# Execution Instructions

Example execution of a standalone script
~~~~
.\create-eop-fw-objects.ps1
~~~~