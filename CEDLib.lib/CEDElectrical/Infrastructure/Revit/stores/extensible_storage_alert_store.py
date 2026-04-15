# -*- coding: utf-8 -*-
"""Extensible storage alert store placeholder."""

from CEDElectrical.Application.contracts.alert_store import IAlertStore


class ExtensibleStorageAlertStore(IAlertStore):
    """Planned storage backend for persistent circuit calculation alerts."""

    def read_alert_payload(self, circuit):
        """Read payload from extensible storage (not implemented yet)."""
        raise NotImplementedError('ExtensibleStorageAlertStore is not implemented yet.')

    def write_alert_payload(self, circuit, payload):
        """Persist payload in extensible storage (not implemented yet)."""
        raise NotImplementedError('ExtensibleStorageAlertStore is not implemented yet.')

    def clear_alert_payload(self, circuit):
        """Clear payload from extensible storage (not implemented yet)."""
        raise NotImplementedError('ExtensibleStorageAlertStore is not implemented yet.')
