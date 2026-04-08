import sys
import argparse
import logging
from pathlib import Path
from typing import Dict, Any

# Fix module search path to allow imports from project root
project_root = Path(__file__).resolve().parent.parent.parent
if str(project_root) not in sys.path:
    sys.path.append(str(project_root))

from core.blueprint_generator.generator import BlueprintGenerator
from core.logger_config import setup_logging
from core.system_config import sys_config

# Use centralized logger
logger = logging.getLogger(__name__)

def main():
    parser = argparse.ArgumentParser(description="Invoice Configuration Generator CLI")
    parser.add_argument("template_path", help="Path to the Excel template file")
    parser.add_argument("-o", "--output", help="Output directory for configuration")
    parser.add_argument("-v", "--verbose", action="store_true", help="Verbose output")
    parser.add_argument("-q", "--quiet", action="store_true", help="Minimal output")
    parser.add_argument("--analyze-only", action="store_true", help="Analyze template and output JSON (legacy format)")
    parser.add_argument("--dry-run", action="store_true", help="Print config but don't save")
    parser.add_argument("--prefix", help="Custom prefix/customer code to use (overrides detection)")

    args = parser.parse_args()

    # Configure logging using centralized logger
    if args.verbose:
        log_level = logging.DEBUG
    elif args.quiet:
        log_level = logging.WARNING
    else:
        log_level = logging.INFO
    
    setup_logging(log_dir=sys_config.run_log_dir, level=log_level)

    template_path = Path(args.template_path).resolve()
    if not template_path.exists():
        logger.error(f"Input file not found: {args.template_path}")
        sys.exit(1)

    try:
        generator = BlueprintGenerator()
        
        if args.analyze_only:
            logger.info(f"Analyzing template: {template_path.name}")
            json_output = generator.analyze(str(template_path))
            
            if args.output:
                output_path = Path(args.output)
                output_path.parent.mkdir(parents=True, exist_ok=True)
                with open(output_path, 'w', encoding='utf-8') as f:
                    f.write(json_output)
                logger.info(f"Analysis saved to: {output_path}")
            else:
                print(json_output)
            sys.exit(0)

        # Standard Generation Flow
        logger.info(f"Starting configuration workflow for: {template_path.name}")
        
        result_path = generator.generate(
            template_path=str(template_path),
            output_dir=args.output,
            dry_run=args.dry_run,
            custom_prefix=args.prefix
        )

        if result_path:
            logger.info(f"Configuration generated successfully at: {result_path}")
            sys.exit(0)
        else:
            logger.error("Generation failed to produce a result path.")
            sys.exit(1)

    except Exception as e:
        logger.error(f"Fatal error: {e}")
        if args.verbose:
            import traceback
            traceback.print_exc()
        sys.exit(1)

if __name__ == "__main__":
    main()
