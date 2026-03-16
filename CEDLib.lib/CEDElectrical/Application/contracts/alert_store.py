# -*- coding: utf-8 -*-
"""Alert persistence contract."""


class IAlertStore(object):
    """Stores and clears persisted alert payloads for a circuit."""

    def read_alert_payload(self, circuit):
        """Read persisted payload for the given circuit."""
        raise NotImplementedError

    def write_alert_payload(self, circuit, payload):
        """Persist payload for the given circuit."""
        raise NotImplementedError

    def clear_alert_payload(self, circuit):
        """Clear persisted payload for the given circuit."""
        raise NotImplementedError
