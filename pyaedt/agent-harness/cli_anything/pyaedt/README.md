# PyAEDT CLI

A stateful CLI harness for [Ansys PyAEDT](https://aedt.docs.pyansys.com/) - Python interface to Ansys Electronics Desktop.

## Overview

This CLI provides command-line access to PyAEDT functionality, enabling AI agents and developers to automate AEDT workflows without a GUI. It supports both one-shot commands and an interactive REPL mode.

## Requirements

- **Windows** - AEDT is Windows-only
- **Ansys Electronics Desktop** - Must be installed with a valid license
- **Python 3.10+**

## Installation

```bash
# Install from source
cd agent-harness
pip install -e .

# Install PyAEDT dependency
pip install pyaedt>=0.6.0
```

## Quick Start

```bash
# Show help
cli-anything-pyaedt --help

# Create a new project
cli-anything-pyaedt project new MyProject --type HFSS

# Set variables
cli-anything-pyaedt variable set freq 10GHz
cli-anything-pyaedt variable set power 1W --scope project

# Create geometry
cli-anything-pyaedt modeler create-box Box1 --position 0,0,0 --dimensions 1,1,1 --material copper

# Create analysis setup
cli-anything-pyaedt analysis create-setup Setup1 --type Hfss --frequency 10GHz

# Output as JSON (for agents)
cli-anything-pyaedt --json project info

# Enter REPL mode
cli-anything-pyaedt
```

## Command Groups

### Project Management
- `project new <name>` - Create new AEDT project
- `project open <file>` - Open existing project
- `project save` - Save current project
- `project close` - Close current project
- `project info` - Show project information

### Design Operations
- `design new <name> --type HFSS` - Create new design
- `design list` - List all designs
- `design activate <name>` - Set active design

### Variable Management
- `variable list` - List all variables
- `variable set <name> <value>` - Set variable value
- `variable get <name>` - Get variable value

### Modeler Operations
- `modeler create-box <name> --position 0,0,0 --dimensions 1,1,1` - Create box
- `modeler create-cylinder <name> --radius 1 --height 10` - Create cylinder
- `modeler assign-material <object> <material>` - Assign material

### Analysis Operations
- `analysis create-setup <name> --type Hfss` - Create analysis setup
- `analysis list` - List all setups
- `analysis run <name>` - Run analysis setup

### Export Operations
- `export results <path> --format csv` - Export results
- `export image <path>` - Export project image
- `export formats` - List supported formats

### Session Management
- `session start --version 2025.1` - Start AEDT session
- `session attach --port 50000` - Attach to running AEDT
- `session status` - Show session status

## JSON Output

All commands support `--json` flag for machine-readable output:

```bash
cli-anything-pyaedt --json project info
```

## REPL Mode

Run without arguments to enter interactive mode:

```bash
cli-anything-pyaedt
```

## AEDT Compatibility

| PyAEDT CLI Version | AEDT Version |
|-------------------|---------------|
| 1.0.0 | 2022 R1+ |

## License

MIT License - See [PyAEDT License](https://github.com/ansys/pyaedt/blob/main/LICENSE)
