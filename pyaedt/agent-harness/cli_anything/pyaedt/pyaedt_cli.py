#!/usr/bin/env python3
"""PyAEDT CLI - Command-line interface for Ansys PyAEDT.

A stateful CLI harness for PyAEDT (Python interface to Ansys Electronics Desktop).
Requires AEDT to be installed (Windows only).

Usage:
    cli-anything-pyaedt --help
    cli-anything-pyaedt project new MyProject
    cli-anything-pyaedt design new HFSSDesign1 --type HFSS
    cli-anything-pyaedt variable set freq 10GHz
    cli-anything-pyaedt modeler create box Box1 --position 0,0,0 --dimensions 1,1,1
    cli-anything-pyaedt --json project info
    cli-anything-pyaedt  # Enter REPL mode
"""

import json
import os
import sys
from pathlib import Path

import click

# Add parent directory to path for development
sys.path.insert(0, str(Path(__file__).parent.parent.parent.parent))

try:
    from cli_anything.pyaedt.utils.repl_skin import ReplSkin
except ImportError:
    # Fallback for direct execution
    from utils.repl_skin import ReplSkin

# Import core modules
from cli_anything.pyaedt.core import project as project_core
from cli_anything.pyaedt.core import design as design_core
from cli_anything.pyaedt.core import variable as variable_core
from cli_anything.pyaedt.core import modeler as modeler_core
from cli_anything.pyaedt.core import analysis as analysis_core
from cli_anything.pyaedt.core import export as export_core
from cli_anything.pyaedt.core import session as session_core


__version__ = "1.0.0"


# Global JSON output flag
json_output = False


def echo_json(data: dict) -> None:
    """Output data as JSON."""
    click.echo(json.dumps(data, indent=2))


def echo_result(result: dict) -> None:
    """Output result based on JSON flag."""
    if json_output:
        echo_json(result)
    elif result.get("status") == "error":
        click.secho(f"✗ {result.get('message', 'Error')}", fg="red", err=True)
    elif result.get("status") == "info":
        click.echo(f"  {result.get('message', '')}")
    else:
        click.secho(f"✓ {result.get('message', 'Success')}", fg="green")


# ── Main CLI Group ────────────────────────────────────────────────────


@click.group(invoke_without_command=True)
@click.option("--json", "use_json", is_flag=True, help="Output results as JSON")
@click.option("--version", is_flag=True, help="Show PyAEDT CLI version")
@click.pass_context
def cli(ctx, use_json, version):
    """PyAEDT CLI - Command-line interface for Ansys PyAEDT.

    Requires AEDT to be installed (Windows only).
    """
    global json_output
    json_output = use_json

    if version:
        echo_json({"version": __version__, "name": "cli-anything-pyaedt"})
        ctx.exit()

    if ctx.invoked_subcommand is None:
        # Enter REPL mode
        ctx.invoke(repl)


# ── REPL Mode ────────────────────────────────────────────────────────


@click.command(name="repl", hidden=True)
@click.pass_context
def repl(ctx):
    """Start interactive REPL mode."""
    skin = ReplSkin("pyaedt", version=__version__)
    skin.print_banner()

    session = session_core.SessionManager()
    status = session.get_status()

    if status["status"] == "disconnected":
        skin.info("Not connected to AEDT")
        skin.hint("Use 'session start' to start AEDT or 'session attach' to connect")

    skin.help({
        "project new <name>": "Create new project",
        "project open <file>": "Open project",
        "project info": "Show project info",
        "design new <name>": "Create new design",
        "design list": "List designs",
        "variable set <name> <value>": "Set variable",
        "variable list": "List variables",
        "modeler create box <name>": "Create box",
        "analysis setup <name>": "Create setup",
        "session start": "Start AEDT session",
        "session status": "Show session status",
        "help": "Show this help",
        "quit": "Exit REPL",
    })

    pt_session = skin.create_prompt_session()

    while True:
        try:
            cmd = skin.get_input(pt_session, context="REPL")
            if not cmd:
                continue

            cmd = cmd.strip().lower()

            if cmd in ["quit", "exit", "q"]:
                break

            if cmd == "help":
                skin.help({
                    "project new <name>": "Create new project",
                    "project open <file>": "Open project",
                    "project info": "Show project info",
                    "design new <name>": "Create new design",
                    "design list": "List designs",
                    "variable set <name> <value>": "Set variable",
                    "variable list": "List variables",
                    "modeler create box <name>": "Create box",
                    "analysis setup <name>": "Create setup",
                    "session start": "Start AEDT session",
                    "session status": "Show session status",
                    "help": "Show this help",
                    "quit": "Exit REPL",
                })
                continue

            if cmd == "session status":
                status = session.get_status()
                skin.status_block(status)
                continue

            skin.warning(f"Unknown command: {cmd}")
            skin.hint("Type 'help' for available commands")

        except KeyboardInterrupt:
            skin.warning("Use 'quit' to exit")
        except EOFError:
            break

    skin.print_goodbye()


# ── Project Commands ──────────────────────────────────────────────────


@cli.group(name="project")
def project():
    """Project management commands."""
    pass


@project.command(name="new")
@click.argument("name")
@click.option("--path", "-p", help="Project directory")
@click.option("--type", "-t", default="HFSS", help="Project type")
def project_new(name, path, type):
    """Create a new AEDT project."""
    result = project_core.create_project(name, path, type)
    echo_result(result)


@project.command(name="open")
@click.argument("project_file")
def project_open(project_file):
    """Open an existing AEDT project."""
    result = project_core.open_project(project_file)
    echo_result(result)


@project.command(name="save")
@click.option("--file", "-f", help="Project file path")
def project_save(file):
    """Save the current project."""
    result = project_core.save_project(file)
    echo_result(result)


@project.command(name="close")
def project_close():
    """Close the current project."""
    result = project_core.close_project()
    echo_result(result)


@project.command(name="info")
@click.option("--file", "-f", help="Project file path")
def project_info(file):
    """Show project information."""
    result = project_core.get_project_info(file)
    if json_output:
        echo_json(result)
    else:
        if result.get("status") == "no_project":
            click.echo("  No project is currently open")
        else:
            click.echo(f"  Name: {result.get('name', 'unknown')}")
            click.echo(f"  Path: {result.get('path', 'unknown')}")


@project.command(name="list")
@click.argument("project_file")
def project_list_designs(project_file):
    """List designs in a project."""
    result = project_core.list_designs(project_file)
    if json_output:
        echo_json(result)
    else:
        if result.get("designs"):
            for d in result["designs"]:
                click.echo(f"  - {d['name']} ({d['type']})")
        else:
            click.echo("  No designs found")


# ── Design Commands ──────────────────────────────────────────────────


@cli.group(name="design")
def design():
    """Design operations commands."""
    pass


@design.command(name="new")
@click.argument("design_name")
@click.option("--type", "-t", default="HFSS", help="Design type")
def design_new(design_name, type):
    """Create a new design."""
    result = design_core.create_design(design_name, type)
    echo_result(result)


@design.command(name="list")
def design_list():
    """List all designs in current project."""
    result = design_core.list_designs()
    echo_result(result)


@design.command(name="activate")
@click.argument("design_name")
def design_activate(design_name):
    """Set the active design."""
    result = design_core.activate_design(design_name)
    echo_result(result)


@design.command(name="info")
@click.argument("design_name", required=False)
def design_info(design_name):
    """Get design information."""
    result = design_core.get_design_info(design_name)
    if json_output:
        echo_json(result)
    else:
        click.echo(f"  Design: {result.get('design_name', 'unknown')}")


# ── Variable Commands ────────────────────────────────────────────────


@cli.group(name="variable")
def variable():
    """Variable management commands."""
    pass


@variable.command(name="list")
@click.option("--scope", "-s", default="all", help="Variable scope")
def variable_list(scope):
    """List variables."""
    result = variable_core.list_variables(scope)
    echo_result(result)


@variable.command(name="set")
@click.argument("name")
@click.argument("value")
@click.option("--scope", "-s", default="design", help="Variable scope")
def variable_set(name, value, scope):
    """Set a variable value."""
    result = variable_core.set_variable(name, value, scope)
    echo_result(result)


@variable.command(name="get")
@click.argument("name")
def variable_get(name):
    """Get a variable value."""
    result = variable_core.get_variable(name)
    echo_result(result)


@variable.command(name="delete")
@click.argument("name")
def variable_delete(name):
    """Delete a variable."""
    result = variable_core.delete_variable(name)
    echo_result(result)


# ── Modeler Commands ──────────────────────────────────────────────────


@cli.group(name="modeler")
def modeler():
    """Modeler operations commands."""
    pass


@modeler.command(name="create-box")
@click.argument("name")
@click.option("--position", "-p", default="0,0,0", help="Position (x,y,z)")
@click.option("--dimensions", "-d", default="1,1,1", help="Dimensions (dx,dy,dz)")
@click.option("--material", "-m", default="aluminum", help="Material name")
def modeler_create_box(name, position, dimensions, material):
    """Create a box in the modeler."""
    pos = [float(x.strip()) for x in position.split(",")]
    dims = [float(x.strip()) for x in dimensions.split(",")]
    result = modeler_core.create_box(name, pos, dims, material)
    echo_result(result)


@modeler.command(name="create-cylinder")
@click.argument("name")
@click.option("--position", "-p", default="0,0,0", help="Position (x,y,z)")
@click.option("--radius", "-r", default=1.0, type=float, help="Radius")
@click.option("--height", "-h", default=10.0, type=float, help="Height")
@click.option("--material", "-m", default="aluminum", help="Material name")
@click.option("--axis", "-a", default="Z", help="Axis (X, Y, or Z)")
def modeler_create_cylinder(name, position, radius, height, material, axis):
    """Create a cylinder in the modeler."""
    pos = [float(x.strip()) for x in position.split(",")]
    result = modeler_core.create_cylinder(name, pos, radius, height, material, axis)
    echo_result(result)


@modeler.command(name="create-sphere")
@click.argument("name")
@click.option("--position", "-p", default="0,0,0", help="Position (x,y,z)")
@click.option("--radius", "-r", default=1.0, type=float, help="Radius")
@click.option("--material", "-m", default="aluminum", help="Material name")
def modeler_create_sphere(name, position, radius, material):
    """Create a sphere in the modeler."""
    pos = [float(x.strip()) for x in position.split(",")]
    result = modeler_core.create_sphere(name, pos, radius, material)
    echo_result(result)


@modeler.command(name="assign-material")
@click.argument("object_name")
@click.argument("material")
def modeler_assign_material(object_name, material):
    """Assign material to an object."""
    result = modeler_core.assign_material(object_name, material)
    echo_result(result)


@modeler.command(name="list")
def modeler_list():
    """List all objects in the modeler."""
    result = modeler_core.list_objects()
    echo_result(result)


@modeler.command(name="info")
@click.argument("object_name")
def modeler_info(object_name):
    """Get object information."""
    result = modeler_core.get_object_info(object_name)
    echo_result(result)


# ── Analysis Commands ──────────────────────────────────────────────────


@cli.group(name="analysis")
def analysis():
    """Analysis operations commands."""
    pass


@analysis.command(name="create-setup")
@click.argument("setup_name")
@click.option("--type", "-t", default="Hfss", help="Setup type")
@click.option("--frequency", "-f", help="Frequency")
def analysis_create_setup(setup_name, type, frequency):
    """Create an analysis setup."""
    props = {"frequency": frequency} if frequency else None
    result = analysis_core.create_setup(setup_name, type, props)
    echo_result(result)


@analysis.command(name="list")
def analysis_list():
    """List all analysis setups."""
    result = analysis_core.list_setups()
    echo_result(result)


@analysis.command(name="run")
@click.argument("setup_name")
def analysis_run(setup_name):
    """Run an analysis setup."""
    result = analysis_core.run_setup(setup_name)
    echo_result(result)


# ── Export Commands ───────────────────────────────────────────────────


@cli.group(name="export")
def export_group():
    """Export operations commands."""
    pass


@export_group.command(name="results")
@click.argument("output_path")
@click.option("--format", "-f", default="csv", help="Export format")
@click.option("--setup", "-s", help="Setup name")
def export_results(output_path, format, setup):
    """Export simulation results."""
    result = export_core.export_results(output_path, format, setup)
    echo_result(result)


@export_group.command(name="image")
@click.argument("output_path")
@click.option("--width", "-w", default=1920, type=int, help="Image width")
@click.option("--height", "-h", default=1080, type=int, help="Image height")
def export_image(output_path, width, height):
    """Export project image."""
    result = export_core.export_image(output_path, width, height)
    echo_result(result)


@export_group.command(name="formats")
def export_formats():
    """List supported export formats."""
    result = export_core.get_export_formats()
    echo_json(result)


# ── Session Commands ──────────────────────────────────────────────────


@cli.group(name="session")
def session():
    """Session management commands."""
    pass


@session.command(name="start")
@click.option("--version", "-v", default="2025.1", help="AEDT version")
@click.option("--non-graphical", "--ng", is_flag=True, help="Non-graphical mode")
@click.option("--port", "-p", default=0, type=int, help="gRPC port")
def session_start(version, non_graphical, port):
    """Start a new AEDT session."""
    result = session_core.start_session(version, non_graphical, port)
    echo_result(result)


@session.command(name="attach")
@click.option("--port", "-p", type=int, required=True, help="gRPC port")
def session_attach(port):
    """Attach to an existing AEDT session."""
    result = session_core.attach_session(port)
    echo_result(result)


@session.command(name="detach")
def session_detach():
    """Detach from the current session."""
    result = session_core.detach_session()
    echo_result(result)


@session.command(name="status")
def session_status():
    """Show current session status."""
    result = session_core.get_session_status()
    if json_output:
        echo_json(result)
    else:
        if result.get("status") == "connected":
            click.echo(f"  Connected to AEDT on port {result.get('port')}")
            click.echo(f"  Version: {result.get('version')}")
        else:
            click.echo("  Not connected to AEDT")


@session.command(name="check")
def session_check():
    """Check if AEDT is available."""
    result = session_core.is_aedt_available()
    echo_result(result)


# ── Process Commands ──────────────────────────────────────────────────


@cli.group(name="process")
def process():
    """Process management commands."""
    pass


@process.command(name="list")
def process_list():
    """List running AEDT processes."""
    # Use PyAEDT's existing CLI functionality
    try:
        import psutil
        aedt_procs = []
        for proc in psutil.process_iter():
            try:
                if proc.name().lower() in ("ansysedt.exe", "ansysedt"):
                    aedt_procs.append({
                        "pid": proc.pid,
                        "name": proc.name(),
                        "status": str(proc.status()),
                    })
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                pass

        if json_output:
            echo_json({"status": "info", "processes": aedt_procs})
        else:
            if aedt_procs:
                for p in aedt_procs:
                    click.echo(f"  PID {p['pid']}: {p['name']} ({p['status']})")
            else:
                click.echo("  No AEDT processes running")
    except ImportError:
        click.secho("  psutil not installed", fg="yellow")


# ── Main Entry Point ──────────────────────────────────────────────────


def main():
    """Main entry point."""
    cli(obj={})


if __name__ == "__main__":
    main()
