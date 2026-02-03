
import sys
import argparse
import logging
from pathlib import Path

# Fix module search path to allow imports from project root
project_root = Path(__file__).resolve().parent.parent.parent
sys.path.append(str(project_root))
sys.path.append(str(Path(__file__).resolve().parent))

try:
    from orchestrator import ConfigOrchestrator
except ImportError:
    # Fallback/Alternative for different run contexts
    from core.blueprint_generator.orchestrator import ConfigOrchestrator

# Use centralized logger - no basicConfig here
logger = logging.getLogger(__name__)

def main():
    parser = argparse.ArgumentParser(description="Invoice Configuration Generator CLI")
    parser.add_argument("template_path", help="Path to the Excel template file")
    parser.add_argument("-o", "--output", help="Output directory for configuration")
    parser.add_argument("-v", "--verbose", action="store_true", help="Verbose output")
    parser.add_argument("--analyze-only", action="store_true", help="Analyze template and output JSON (legacy format)")

    args = parser.parse_args()

    if args.verbose:
        logger.setLevel(logging.DEBUG)

    template_path = Path(args.template_path)
    base_dir = Path(__file__).resolve().parent

    try:
        orchestrator = ConfigOrchestrator(base_dir)
        options = {}
        if args.output:
            options['output_dir'] = args.output
        if args.analyze_only:
            options['analyze_only'] = True

        success = orchestrator.run(str(template_path), options) # Pass str path for compatibility

        if success:
            sys.exit(0)
        else:
            sys.exit(1)

    except Exception as e:
        logger.error(f"Fatal error: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
