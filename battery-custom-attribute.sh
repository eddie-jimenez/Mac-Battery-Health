#!/bin/bash
# macOS Battery Health - Intune Custom Attribute (CSV, single line, echo-only)
# Columns:
# HealthPercent,CycleCount,FullChargeCapacity_mAh,DesignCapacity_mAh,CurrentCapacity_mAh,IsCharging,ExternalPowerConnected,TimeRemainingMin,Voltage_mV,Condition,OverThreshold

PB="/usr/libexec/PlistBuddy"
TMP="$(/usr/bin/mktemp -t batt).plist"
trap 'rm -f "$TMP"' EXIT

# Read AppleSmartBattery as a plist with error handling
if ! /usr/sbin/ioreg -a -r -c AppleSmartBattery > "$TMP" 2>/dev/null; then
  echo "None"
  exit 0
fi

# Verify we have valid data
if ! /usr/bin/grep -q "<dict>" "$TMP" 2>/dev/null; then
  echo "None"
  exit 0
fi

# Safe PlistBuddy reader
pb() { 
  "$PB" -c "Print :0:$1" "$TMP" 2>/dev/null || echo ""
}

# Value converters with better error handling
intval() { 
  local val="$1"
  # Remove any non-numeric characters except minus sign
  val="${val//[^0-9-]/}"
  # Check if it's a valid integer
  if [[ "$val" =~ ^-?[0-9]+$ ]]; then
    echo "$val"
  else
    echo "0"
  fi
}

yn() {
  local v="$(echo "$1" | /usr/bin/tr '[:upper:]' '[:lower:]')"
  case "$v" in 
    yes|true|1) echo "True" ;; 
    no|false|0) echo "False" ;; 
    *) echo "None" ;; 
  esac
}

val_or_none() { 
  local val="$1"
  if [[ "$val" =~ ^-?[0-9]+$ ]] && [ "$val" != "0" ]; then
    echo "$val"
  else
    echo "None"
  fi
}

# Raw reads with fallbacks
design_s="$(pb DesignCapacity)"
rawmax_s="$(pb AppleRawMaxCapacity)"
nominal_s="$(pb NominalChargeCapacity)"
max_s="$(pb MaxCapacity)"
cur_s="$(pb CurrentCapacity)"
rawcur_s="$(pb AppleRawCurrentCapacity)"
cycles_s="$(pb CycleCount)"
ischg_s="$(pb IsCharging)"
ext_s="$(pb ExternalConnected)"
trem_s="$(pb TimeRemaining)"
volt_s="$(pb Voltage)"

# Get battery condition - try multiple possible field names
cond_s=""
for field in "PermanentFailureStatus" "BatteryHealth" "HealthInfo:Condition" "BatteryHealthCondition"; do
  test_val="$(pb "$field")"
  if [ -n "$test_val" ] && [ "$test_val" != "0" ]; then
    cond_s="$test_val"
    break
  fi
done

# If still no condition, check system_profiler as fallback
if [ -z "$cond_s" ] || [ "$cond_s" = "0" ]; then
  cond_s="$(/usr/sbin/system_profiler SPPowerDataType 2>/dev/null | /usr/bin/grep -E "Condition:|Health Information:" | /usr/bin/head -1 | /usr/bin/sed 's/.*: *//')"
fi

# Map numeric condition values to text if needed
case "$cond_s" in
  "0"|"Good"|"Normal") cond_s="Normal" ;;
  "1"|"Fair") cond_s="Replace Soon" ;;
  "2"|"Poor") cond_s="Replace Now" ;;
  "3"|"Check Battery") cond_s="Service Battery" ;;
  "") cond_s="Normal" ;;  # Default to Normal if unknown
  *) ;; # Keep as-is if it's already text
esac

# Convert to integers with validation
i_design=$(intval "$design_s")
i_rawmax=$(intval "$rawmax_s")
i_nominal=$(intval "$nominal_s")
i_max=$(intval "$max_s")
i_cur=$(intval "$cur_s")
i_rawcur=$(intval "$rawcur_s")
i_cycles=$(intval "$cycles_s")
i_trem=$(intval "$trem_s")
i_volt=$(intval "$volt_s")

# FullChargeCapacity (mAh) - improved logic
fcc=0
if [ "$i_rawmax" -gt 0 ]; then
  fcc=$i_rawmax
elif [ "$i_nominal" -gt 0 ] && [ "$i_max" -gt 0 ] && [ "$i_max" -le 110 ]; then
  # If MaxCapacity is a percentage, calculate from nominal
  fcc=$(( (i_nominal * i_max) / 100 ))
elif [ "$i_max" -gt 200 ]; then
  # If MaxCapacity is in mAh
  fcc=$i_max
fi

# Current capacity (mAh)
curr=0
if [ "$i_rawcur" -gt 0 ]; then
  curr=$i_rawcur
elif [ "$i_cur" -gt 200 ]; then
  curr=$i_cur
elif [ "$i_nominal" -gt 0 ] && [ "$i_cur" -gt 0 ] && [ "$i_cur" -le 100 ]; then
  # If CurrentCapacity is a percentage
  curr=$(( (i_nominal * i_cur) / 100 ))
fi

# Health % - improved calculation
health="None"
if [ "$i_design" -gt 0 ] && [ "$fcc" -gt 0 ]; then
  health=$(( (100 * fcc) / i_design ))
  # Sanity check - battery health shouldn't exceed 110%
  if [ "$health" -gt 110 ]; then
    health=100
  fi
elif [ "$i_max" -gt 0 ] && [ "$i_max" -le 110 ]; then
  health=$i_max
fi

# Time remaining (65535 means calculating/unknown)
trem_out="None"
if [ "$i_trem" -gt 0 ] && [ "$i_trem" -lt 65535 ]; then 
  trem_out="$i_trem"
fi

# Threshold flag
CYCLE_MAX=1000
over="False"
[ "$i_cycles" -ge "$CYCLE_MAX" ] && over="True"

# Build and output CSV line with proper escaping
echo "$(val_or_none "$health"),$(val_or_none "$i_cycles"),$(val_or_none "$fcc"),$(val_or_none "$i_design"),$(val_or_none "$curr"),$(yn "$ischg_s"),$(yn "$ext_s"),$trem_out,$(val_or_none "$i_volt"),$cond_s,$over"

exit 0