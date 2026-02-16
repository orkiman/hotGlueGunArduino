# Hot Glue Gun Arduino Controller

A precision Arduino-based controller for automated hot glue gun operations, designed for synchronized paper processing with dual glue guns and real-time position tracking.

## Overview

This system controls two hot glue guns in synchronization with paper movement through a machine. It uses an encoder for precise position tracking and a photocell sensor to detect paper sheets, enabling accurate glue application according to predefined patterns.

## Hardware Requirements

### Arduino Board
- Any Arduino-compatible board with sufficient EEPROM storage
- Supports ESP8266/ESP32 for additional features

### Pin Configuration
- **PIN_ENCODER_A (2)**: Encoder input for position tracking
- **PIN_PHOTOCELL (3)**: Photocell sensor for paper detection
- **PIN_GUN1 (8)**: Control output for hot glue gun 1
- **PIN_GUN2 (9)**: Control output for hot glue gun 2

### External Components
- Rotary encoder for position feedback
- Photocell/optical sensor for paper edge detection
- Two hot glue guns with relay/solid-state control
- Power supply for glue guns (separate from Arduino)

## Features

### Core Functionality
- **Dual Gun Control**: Independent control of two hot glue guns
- **Position-Based Operation**: Glue application based on precise position tracking
- **Pattern Programming**: Define complex glue patterns with line segments
- **Multi-Sheet Handling**: Track and process up to 10 sheets simultaneously
- **Speed Monitoring**: Automatic shutdown if paper moves too slowly

### Configuration Management
- **Persistent Storage**: All settings saved to EEPROM
- **CRC Validation**: Data integrity verification
- **Default Fallbacks**: Safe operation if configuration corrupted

### Calibration & Testing
- **Automatic Calibration**: Measure encoder pulses per millimeter
- **Manual Testing**: Override controls for gun testing
- **Real-time Feedback**: Serial communication for monitoring

## Data Structures

### Configuration Parameters
```cpp
struct Config {
  float pulses_per_mm;        // Encoder resolution (pulses/mm)
  uint16_t max_ms_per_mm;     // Maximum speed threshold (ms/mm)
  float photocell_offset_mm;  // Sensor offset from reference point
  uint16_t debounce_ms;       // Photocell debounce time
};
```

### Pattern Definition
```cpp
struct Line {
  float start_mm;  // Line start position
  float end_mm;    // Line end position
};

struct Pattern {
  Line lines[MAX_LINES_PER_GUN];  // Up to 32 lines per gun
  uint8_t count;                  // Number of lines in pattern
};
```

### Sheet Tracking
```cpp
struct Sheet {
  bool active;                   // Sheet is currently being processed
  float mm;                      // Current position
  uint32_t started_ms;           // Start timestamp
  int32_t last_mm_int;           // Last integer position
  uint32_t last_mm_change_ms;    // Last position change time
  bool slow_block[2];            // Speed violation flags for each gun
};
```

## Command Interface

The system accepts NDJSON (Newline Delimited JSON) commands via Serial at 115200 baud.

### Control Commands

#### Activate/Deactivate System
```json
{"cmd":"set_active","active":true}
```

#### Configure System Parameters
```json
{"cmd":"set_config","pulses_per_mm":12.34,"max_ms_per_mm":80,"photocell_offset_mm":250.0,"debounce_ms":20}
```

#### Set Glue Patterns
```json
{"cmd":"set_pattern","gun":1,"lines":[{"start":10,"end":40},{"start":60,"end":90}]}
{"cmd":"set_pattern","gun":2,"lines":[{"start":15,"end":35}]}
```

#### Calibration
```json
{"cmd":"calib_arm","paper_length_mm":297.0}
```

#### Testing Commands
```json
{"cmd":"test_open","gun":1,"timeout_ms":2000}
{"cmd":"test_close","gun":1}
{"cmd":"test_open","gun":"both","timeout_ms":1000}
{"cmd":"test_close","gun":"both"}
```

## Operation Modes

### Normal Operation
1. System activated via `set_active` command
2. Paper sheets detected by photocell sensor
3. Position tracking starts from photocell trigger point
4. Glue guns activate based on pattern matching current position
5. Sheets automatically removed when pattern complete or timeout reached

### Calibration Mode
1. Send `calib_arm` command with known paper length
2. Pass paper through machine twice
3. System automatically calculates and saves `pulses_per_mm`
4. Calibration result sent via Serial

### Test Mode
- Manual gun control for testing and maintenance
- Timeout-based automatic shutoff
- Supports individual or simultaneous gun control

## Safety Features

### Speed Monitoring
- Tracks paper movement speed in real-time
- Disables glue guns if movement too slow (`max_ms_per_mm` threshold)
- Prevents glue buildup from slow movement

### Timeout Protection
- 30-second sheet timeout prevents infinite tracking
- Test commands have configurable timeouts
- Automatic system shutdown on inactivity

### Data Integrity
- CRC validation of stored configuration
- Magic number and version checking
- Graceful fallback to defaults on corruption

## Performance Specifications

### Position Tracking
- Resolution: Determined by encoder (typically 10-20 pulses/mm)
- Accuracy: ±0.1mm (depends on encoder quality)
- Update rate: Real-time interrupt-driven

### Pattern Capacity
- Maximum lines per gun: 32
- Maximum simultaneous sheets: 10
- Pattern range: Limited by float precision (±3.4m)

### Timing
- Photocell debounce: 0-1000ms (configurable)
- Speed monitoring: 1-60000ms/mm (configurable)
- Sheet timeout: 30 seconds (fixed)

## Installation & Setup

1. **Hardware Installation**
   - Connect encoder to PIN_ENCODER_A
   - Connect photocell to PIN_PHOTOCELL
   - Connect gun controls to PIN_GUN1/PIN_GUN2
   - Ensure proper power supply separation

2. **Upload Firmware**
   - Open `hotGlueGunArduino.ino` in Arduino IDE
   - Select appropriate board and port
   - Upload firmware

3. **Initial Configuration**
   - Connect Serial monitor at 115200 baud
   - Set basic configuration parameters
   - Calibrate encoder using known paper length

4. **Pattern Programming**
   - Define glue patterns for each gun
   - Test with manual commands
   - Activate system for production

## Desktop Configuration GUI

A complete Windows desktop configuration program is included in:

- `ConfigApp/GlueConfigApp.csproj`

### What it provides

- COM port connect/disconnect and live serial log
- Full config editing (`pulses_per_mm`, `max_ms_per_mm`, `photocell_offset_mm`, `debounce_ms`)
- Pattern editor for each gun (start/end intervals)
- Small visual pattern preview per gun
- One-click send commands (`set_config`, `set_pattern`, `set_active`)
- Test control buttons (`test_open`/`test_close`) for gun 1, gun 2, and both guns
- Calibration controls (`paper_length_mm` + `calib_arm` command)
- Validation rules before send/save (line count, allowed range, no overlapping intervals)
- Save and load glue programs as JSON files

### Run the GUI

From the project root:

```powershell
dotnet run --project .\ConfigApp\GlueConfigApp.csproj
```

### Build the GUI

```powershell
dotnet build .\ConfigApp\GlueConfigApp.csproj
```

The project targets `net10.0-windows`, so install .NET 10 SDK before build/run.

### Glue program file contents

Saved glue program JSON includes:

- `config`
- `gun1Lines`
- `gun2Lines`

Load a saved file, then click **Send All** to push configuration and patterns to the controller.

## Troubleshooting

### Common Issues

**Glue guns not activating**
- Check system is active (`set_active:true`)
- Verify pattern configuration
- Check speed monitoring isn't blocking
- Test with manual override commands

**Inaccurate positioning**
- Recalibrate encoder with known paper length
- Check encoder connections
- Verify `pulses_per_mm` configuration

**False paper detection**
- Adjust photocell debounce timing
- Check sensor alignment
- Verify photocell offset configuration

**Configuration lost**
- Check EEPROM size compatibility
- Verify CRC validation
- Reconfigure with default values

### Debug Mode
Set `DEBUG` to `1` at top of file to enable Serial debug output for troubleshooting.

## Development Notes

### Memory Usage
- EEPROM: ~200 bytes for persistent storage
- RAM: ~500 bytes for runtime data structures
- Stack: Minimal, designed for real-time operation

### Interrupt Handling
- Encoder pulses counted via hardware interrupt
- Critical sections protected with interrupt disable
- Non-blocking design for reliable operation

### Extensibility
- Easy to add new commands via `handleJsonLine()`
- Pattern system supports complex geometries
- Modular design for feature additions

## License

This project is provided as-is for educational and industrial use. Modify as needed for specific applications.
