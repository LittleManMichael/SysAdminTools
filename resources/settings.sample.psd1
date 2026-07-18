#
# Settings.sample.psd1 - TEMPLATE ONLY. Placeholder values.
#
# Each network keeps its own real copy of this file, named Settings.psd1, in
# the Resources folder of its authoritative share. 
#
# Same filename and keys on every network; only the values differ. This is
# what lets every script stay byte-identical across all three networks:
# scripts read these values through Get-SysAdminToolsSetting instead of
# hardcoding paths, relays, or name patterns.
#
# Keys grow over time as commands are ported. When a script needs a new
# environment value, add the key HERE first (with a placeholder), then to
# each network's real Settings.psd1.
#
# FYI: The ` is to continue on the next line. Allows code to be written
# and seen cleaner

@{

    # Short name for this network. Shown by Get-SysAdminToolsInfo so an admin
    # can always confirm which network's settings are loaded.
    NetworkName = 'SAMPLE'

    # Authoritative SysAdminTools root on this network's admin share.
    # Reports, Logs, and Archive resolve underneath this path.
    ShareRoot = '\\server\share\SysAdminTools'

    # Mail
    SmtpRelay  = 'relay.example.local'
    MailDomain = 'example.local'

    DistributionLists = `
    @{
        SysAdmins = '-sysadmins@example.local'
    }

    # Active Directory
    DomainSuffix = 'example.local'

    # Server name patterns used by AD discovery filters
    # (e.g. drive space checks, server availability sweeps).
    ServerNamePatterns = `
    @{
        DomainControllers = @('*DC0*')
        FilePrint         = @('*FP0*')
        Exchange          = @('*MBX*')
        Exclude           = @('*XDC*')
    }

}
