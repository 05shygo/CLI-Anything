---
name: cli-anything-pyaedt
description: CLI harness for Ansys PyAEDT - Python interface to Ansys Electronics Desktop
version: 1.0.0
command_groups:
  project:
    name: project
    description: Project management commands
    commands:
      - name: project new
        args:
          - name: name
            type: string
            required: true
            description: Project name
          - name: --path
            type: string
            required: false
            description: Project directory
          - name: --type
            type: string
            required: false
            default: HFSS
            description: Project type (HFSS, Maxwell3D, Icepak, Circuit)
        description: Create new AEDT project
      - name: project open
        args:
          - name: project_file
            type: string
            required: true
            description: Path to .aedt project file
        description: Open existing AEDT project
      - name: project save
        args:
          - name: --file
            type: string
            required: false
            description: Project file path
        description: Save current project
      - name: project close
        description: Close current project
      - name: project info
        args:
          - name: --file
            type: string
            required: false
            description: Project file path
        description: Show project information
  design:
    name: design
    description: Design operations
    commands:
      - name: design new
        args:
          - name: design_name
            type: string
            required: true
            description: Design name
          - name: --type
            type: string
            required: false
            default: HFSS
            description: Design type
        description: Create new design
      - name: design list
        description: List all designs
      - name: design activate
        args:
          - name: design_name
            type: string
            required: true
            description: Design name
        description: Set active design
  variable:
    name: variable
    description: Variable management
    commands:
      - name: variable list
        args:
          - name: --scope
            type: string
            required: false
            default: all
            description: Variable scope (all, project, design)
        description: List variables
      - name: variable set
        args:
          - name: name
            type: string
            required: true
            description: Variable name
          - name: value
            type: string
            required: true
            description: Variable value with units
          - name: --scope
            type: string
            required: false
            default: design
            description: Variable scope
        description: Set variable value
      - name: variable get
        args:
          - name: name
            type: string
            required: true
            description: Variable name
        description: Get variable value
  modeler:
    name: modeler
    description: Geometry operations
    commands:
      - name: modeler create-box
        args:
          - name: name
            type: string
            required: true
            description: Box name
          - name: --position
            type: string
            required: false
            default: "0,0,0"
            description: Position (x,y,z)
          - name: --dimensions
            type: string
            required: false
            default: "1,1,1"
            description: Dimensions (dx,dy,dz)
          - name: --material
            type: string
            required: false
            default: aluminum
            description: Material name
        description: Create box
      - name: modeler create-cylinder
        args:
          - name: name
            type: string
            required: true
            description: Cylinder name
          - name: --position
            type: string
            required: false
            default: "0,0,0"
            description: Position (x,y,z)
          - name: --radius
            type: float
            required: false
            default: 1.0
            description: Radius
          - name: --height
            type: float
            required: false
            default: 10.0
            description: Height
          - name: --axis
            type: string
            required: false
            default: Z
            description: Axis (X, Y, Z)
        description: Create cylinder
      - name: modeler assign-material
        args:
          - name: object_name
            type: string
            required: true
            description: Object name
          - name: material
            type: string
            required: true
            description: Material name
        description: Assign material
  analysis:
    name: analysis
    description: Analysis operations
    commands:
      - name: analysis create-setup
        args:
          - name: setup_name
            type: string
            required: true
            description: Setup name
          - name: --type
            type: string
            required: false
            default: Hfss
            description: Setup type
        description: Create analysis setup
      - name: analysis list
        description: List setups
      - name: analysis run
        args:
          - name: setup_name
            type: string
            required: true
            description: Setup name
        description: Run analysis
  export:
    name: export
    description: Export operations
    commands:
      - name: export results
        args:
          - name: output_path
            type: string
            required: true
            description: Output file path
          - name: --format
            type: string
            required: false
            default: csv
            description: Format (csv, hdf5, snp)
        description: Export results
      - name: export image
        args:
          - name: output_path
            type: string
            required: true
            description: Image output path
        description: Export image
      - name: export formats
        description: List formats
  session:
    name: session
    description: Session management
    commands:
      - name: session start
        args:
          - name: --version
            type: string
            required: false
            default: "2025.1"
            description: AEDT version
          - name: --non-graphical
            type: boolean
            required: false
            description: Non-graphical mode
        description: Start AEDT session
      - name: session attach
        args:
          - name: --port
            type: integer
            required: true
            description: gRPC port
        description: Attach to session
      - name: session status
        description: Show status
      - name: session check
        description: Check AEDT availability
  process:
    name: process
    description: Process management
    commands:
      - name: process list
        description: List AEDT processes
---

# PyAEDT CLI Skill

## Overview

This skill provides CLI access to Ansys PyAEDT, a Python interface to Ansys Electronics Desktop (AEDT). AEDT is a Windows-only application for electromagnetic and circuit simulation.

## Requirements

- Windows OS with AEDT installed
- Python 3.10+
- PyAEDT package

## Common Workflows

### Create and Configure HFSS Project

```bash
# Create new HFSS project
cli-anything-pyaedt project new my_project --type HFSS

# Set design variables
cli-anything-pyaedt variable set freq 10GHz
cli-anything-pyaedt variable set power 1W

# Create geometry
cli-anything-pyaedt modeler create-box antenna --position 0,0,0 --dimensions 10,10,0.5 --material aluminum

# Create analysis setup
cli-anything-pyaedt analysis create-setup sim1 --type Hfss --frequency 10GHz
```

### Create Maxwell 3D Design

```bash
# Create Maxwell project
cli-anything-pyaedt project new motor_design --type Maxwell3D

# Create design
cli-anything-pyaedt design new mag_design --type Maxwell3D

# Set project variables
cli-anything-pyaedt variable set $current 100A --scope project

# Create geometry
cli-anything-pyaedt modeler create-cylinder stator --position 0,0,0 --radius 50 --height 20 --axis Z
```

### Export Results

```bash
# Export as CSV
cli-anything-pyaedt export results output.csv --format csv

# Export as touchstone
cli-anything-pyaedt export results s_params.snp --format snp

# Export geometry
cli-anything-pyaedt export results model.step
```

## JSON Output

All commands support `--json` flag for machine-readable output:

```bash
cli-anything-pyaedt --json project info
cli-anything-pyaedt --json variable list
```

## Session Management

```bash
# Start AEDT session
cli-anything-pyaedt session start --version 2025.1 --non-graphical

# Attach to running AEDT
cli-anything-pyaedt session attach --port 50000

# Check status
cli-anything-pyaedt session status
```

## Notes

- AEDT is Windows-only; most commands will not function on Linux/macOS
- Full functionality requires a licensed AEDT installation
- Some operations require an active AEDT session
- Design types include: HFSS, HFSS3DLayout, Maxwell2D, Maxwell3D, Icepak, Q3D, Circuit, EMIT, TwinBuilder
