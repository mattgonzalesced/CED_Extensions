"""Helpers to validate user overrides before calculating circuits."""


class OverrideValidator(object):
    def __init__(self, logger=None):
        self.logger = logger

    def validate(self, branch):
        messages = []

        auto_override = getattr(branch, '_auto_calculate_override', False)
        overrides = {
            'hot': getattr(branch, '_wire_hot_size_override', None),
            'neutral': getattr(branch, '_wire_neutral_size_override', None),
            'ground': getattr(branch, '_wire_ground_size_override', None),
            'sets': getattr(branch, '_wire_sets_override', None),
            'conduit_type': getattr(branch, '_conduit_type_override', None),
            'conduit_size': getattr(branch, '_conduit_size_override', None),
        }

        if auto_override and not any(overrides.values()):
            messages.append(
                "{} has user overrides enabled but no override values were found; calculations will be used.".format(
                    branch.name
                )
            )

        return messages
