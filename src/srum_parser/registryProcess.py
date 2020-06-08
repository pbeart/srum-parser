"Not yet implemented"

from regipy.registry import RegistryHive


class RegistryProcessor:

    def __init__(self, ntuser=None, system=None, security=None, sam=None):
        raise NotImplementedError

        self.ntuser = RegistryHive(ntuser) if ntuser is not None else None
        self.system = RegistryHive(system) if system is not None else None
        self.security = RegistryHive(security) if security is not None else None
        self.sam = RegistryHive(sam) if sam is not None else None

    def interface_luid_to_ssid(self, luid):
        pass
