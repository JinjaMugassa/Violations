"""Violation report processors package.

This package contains individual processors for each type of violation report:
- speed_violation: Speed violation report processor
- harsh_brake: Harsh brake violations processor (with merge capability)
- idling: Idling violations processor
- night_driving: Night driving report processor

Each processor module exports:
- TEMPLATE_ID: The Wialon template ID for the report
- TEMPLATE_NAME: Human-readable template name
- process_*(): Processing function that takes (df, template_id, api)
"""

from .speed_violation import (
    TEMPLATE_ID as SPEED_TEMPLATE_ID,
    TEMPLATE_NAME as SPEED_TEMPLATE_NAME,
    process_speed_violation
)

from .harsh_brake import (
    SUMMARY_TEMPLATE_ID,
    DETAIL_TEMPLATE_ID,
    SUMMARY_TEMPLATE_NAME,
    DETAIL_TEMPLATE_NAME,
    process_harsh_brake_detail,
    merge_harsh_brake_reports
)

from .idling import (
    TEMPLATE_ID as IDLING_TEMPLATE_ID,
    TEMPLATE_NAME as IDLING_TEMPLATE_NAME,
    process_idling
)

from .night_driving import (
    TEMPLATE_ID as NIGHT_TEMPLATE_ID,
    TEMPLATE_NAME as NIGHT_TEMPLATE_NAME,
    process_night_driving
)

__all__ = [
    # Speed violation
    'SPEED_TEMPLATE_ID',
    'SPEED_TEMPLATE_NAME',
    'process_speed_violation',
    
    # Harsh brake
    'SUMMARY_TEMPLATE_ID',
    'DETAIL_TEMPLATE_ID',
    'SUMMARY_TEMPLATE_NAME',
    'DETAIL_TEMPLATE_NAME',
    'process_harsh_brake_detail',
    'merge_harsh_brake_reports',
    
    # Idling
    'IDLING_TEMPLATE_ID',
    'IDLING_TEMPLATE_NAME',
    'process_idling',
    
    # Night driving
    'NIGHT_TEMPLATE_ID',
    'NIGHT_TEMPLATE_NAME',
    'process_night_driving',
]