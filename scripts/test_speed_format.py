import sys
import os
import pandas as pd

# Ensure project root is on sys.path so `processors` package can be imported
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from processors.speed_violation import process_speed_violation

# Sample rows with seconds present
df = pd.DataFrame({
    'Time': ['2026-01-07 10:44:06', '2026-01-07 15:50:08', '2026-01-07 07:51:55'],
    'Speed': ['90 km/h', '86 km/h', '84 km/h']
})

out = process_speed_violation(df, 3, None)
print(out.to_string(index=False))
