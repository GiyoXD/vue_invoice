import logging
import sys
from pathlib import Path
from typing import Dict, Any

try:
    from blueprint_generator import BlueprintGenerator
    from core.workspace import WorkspaceManager
except ImportError:
    from .blueprint_generator import BlueprintGenerator
    from .core.workspace import WorkspaceManager

logger = logging.getLogger(__name__)

class ConfigOrchestrator:
    """
    Facade for the configuration generation process.
    Refactored to directly use BlueprintGenerator, removing redundant analysis steps.
    """

    def __init__(self, base_dir: Path):
        self.base_dir = base_dir
        # WorkspaceManager is still useful for managing output directories if needed,
        # but BlueprintGenerator handles most of it.
        self.workspace = WorkspaceManager(base_dir / "result")
        self.generator = BlueprintGenerator()

    def run(self, excel_file: str, options: Dict[str, Any] = None) -> bool:
        """
        Executes the configuration generation using AutoBundleGenerator.
        """
        if options is None:
            options = {}
            
        excel_path = Path(excel_file).resolve()
        if not excel_path.exists():
            logger.error(f"Input file not found: {excel_file}")
            return False

        logger.info(f"Starting configuration workflow for: {excel_path.name}")

        try:
            # Determine output directory/file
            custom_output = options.get('output_dir')
            
            # Check for analyze_only mode
            if options.get('analyze_only'):
                json_output = self.generator.analyze(str(excel_path))
                
                if custom_output:
                    # Treat custom_output as a file path for analysis JSON
                    output_path = Path(custom_output)
                    output_path.parent.mkdir(parents=True, exist_ok=True)
                    with open(output_path, 'w', encoding='utf-8') as f:
                        f.write(json_output)
                    logger.info(f"Analysis saved to: {output_path}")
                else:
                    # Print to stdout
                    print(json_output)
                return True

            # Standard Generation Flow
            output_dir = Path(custom_output) if custom_output else None

            # Generate Bundle Config
            # AutoBundleGenerator handles analysis and building in one go.
            result_path = self.generator.generate(
                template_path=str(excel_path),
                output_dir=str(output_dir) if output_dir else None,
                dry_run=False
            )

            if result_path:
                logger.info(f"Configuration generated successfully at: {result_path}")
                return True
            else:
                logger.error("Generation failed to produce a result path.")
                return False

        except Exception as e:
            logger.error(f"Orchestration failed: {e}")
            import traceback
            traceback.print_exc()
            return False

