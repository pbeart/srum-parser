"Provide functions to handle processing of certain columns from SRUM tables into a more useful format"


def parse_notification_type(type_id):
    """https://docs.microsoft.com/en-us/uwp/api/windows.networking.pushnotifications.pushnotificationtype?view=winrt-18362
    Returns a named notification type from a notification type ID"""

    notification_types = {
        0: "Toast",
        1: "Tile",
        2: "Badge",
        3: "Raw",
        4: "TileFlyout"
    }

    if type_id in notification_types:
        return notification_types[type_id]

    else:
        return "Unknown notification type"
