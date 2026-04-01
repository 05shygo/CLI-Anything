# PyAEDT CLI Harness - Software-Specific SOP

## Overview

**Software**: Ansys PyAEDT (Python interface to Ansys Electronics Desktop)
**Backend**: PyAEDT Python library communicating with AEDT via gRPC/COM
**Native Format**: AEDT project files (.aedt), EDB databases (.aedb)
**System Dependency**: AEDT (Windows-only, requires license)

## Architecture Analysis

### Backend Engine
- **AEDT (Ansys Electronics Desktop)**: Windows-only GUI application for electromagnetic and circuit simulation
- **PyAEDT**: Python library that provides a Pythonic API to AEDT's functionality
- Communication: gRPC (primary, cross-platform) or COM (Windows-only)

### Key AEDT Applications
| Tool | Purpose | File Format |
|------|---------|-------------|
| HFSS | 3D electromagnetic simulation | .aedt |
| Maxwell | 2D/3D electromagnetic field simulation | .aedt |
| Icepak | Thermal analysis | .aedt |
| Q3D Extractor | Parameter extraction | .aedt |
| Nexxim | Circuit simulation | .aedt |
| EDB | Electronic database for PCB/layout | .aedb |
| EMIT | RF propagation modeling | .aedt |
| Twin Builder | System simulation | .aedt |

### Data Model
- **Project files**: `.aedt` (XML-based, ZIP compressed)
- **EDB databases**: `.aedb` directories containing layout data
- **Designs**: Each project can contain multiple designs
- **Variables**: Project and design-level parameters

## CLI Command Groups

### 1. Project Management (`project`)
- `project new` - Create new AEDT project
- `project open` - Open existing project
- `project save` - Save current project
- `project close` - Close project
- `project info` - Display project information
- `project list` - List designs in project

### 2. Design Operations (`design`)
- `design new` - Create new design in project
- `design list` - List available designs
- `design activate` - Set active design
- `design copy` - Copy design

### 3. Variable Management (`variable`)
- `variable list` - List all variables
- `variable set` - Set variable value
- `variable get` - Get variable value

### 4. Modeler Operations (`modeler`)
- `modeler create box` - Create box geometry
- `modeler create cylinder` - Create cylinder geometry
- `modeler assign material` - Assign material to object
- `modeler list` - List model objects

### 5. Analysis Operations (`analysis`)
- `analysis setup` - Create analysis setup
- `analysis run` - Run analysis
- `analysis list` - List analysis setups

### 6. Export Operations (`export`)
- `export results` - Export simulation results
- `export image` - Export project image

### 7. Session Management (`session`)
- `session start` - Start AEDT session
- `session attach` - Attach to running AEDT
- `session detach` - Detach from AEDT
- `session status` - Show session status

### 8. Process Management (`process`)
- `process list` - List running AEDT processes
- `process start` - Start AEDT process
- `process stop` - Stop AEDT process

## State Model

### Session State
- Current AEDT port (if connected via gRPC)
- Active project path
- Active design name
- Project variable values

### Project State (stored in .json session files)
```json
{
  "project_path": "C:/path/to/project.aedt",
  "design_name": "HFSSDesign1",
  "aedt_version": "2025.1",
  "variables": {"freq": "10GHz", "power": "1W"},
  "port": 50000
}
```

## Output Formats

### Human-Readable (default)
- Formatted tables for lists
- Status messages with icons
- Progress indicators for long operations

### JSON (--json flag)
```json
{
  "status": "success",
  "command": "project-info",
  "data": {
    "name": "MyProject",
    "path": "C:/path/MyProject.aedt",
    "designs": ["HFSSDesign1", "MaxwellDesign1"]
  }
}
```

## Backend Integration

### Finding AEDT
PyAEDT uses `ansysedt.exe` (Windows) or searches installed locations.

### Launching AEDT
```python
from ansys.aedt.core import Desktop
with Desktop(version="2025.1", non_graphical=True) as d:
    hfss = Hfss()
    # operations
```

### Communication
- **gRPC**: Default for cross-platform support
- **COM**: Windows-only, fallback

## Critical Notes

1. **AEDT is Windows-only**: Most operations require Windows with AEDT installed
2. **License required**: AEDT needs a valid license
3. **Non-graphical mode**: Use `--non-graphical` for headless operation
4. **EDB operations**: Can work without full AEDT via PyEDB gRPC
