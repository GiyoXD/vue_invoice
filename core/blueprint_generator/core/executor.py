import subprocess
import sys
import logging

logger = logging.getLogger(__name__)

class Executor:
    """
    Handles safe execution of subprocess commands.
    """

    def run_command(self, command: list, verbose: bool = False) -> bool:
        """
        Executes a command in a subprocess and handles output.

        Args:
            command (list): The command and its arguments to execute.
            verbose (bool): If True, print the command being executed.

        Returns:
            bool: True for success, False for failure.
        """
        if verbose:
            logger.info(f"Running command: {' '.join(command)}")

        try:
            process = subprocess.Popen(
                command,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                encoding='utf-8',
                errors='replace'
            )

            stdout, stderr = process.communicate()

            if process.returncode != 0:
                logger.error(f"Error executing: {' '.join(command)}")
                if stdout:
                    logger.error("--- STDOUT ---")
                    logger.error(stdout)
                if stderr:
                    logger.error("--- STDERR ---")
                    logger.error(stderr)
                return False

            if verbose and stdout:
                print(stdout)

            return True

        except Exception as e:
            logger.error(f"Exception during command execution: {e}")
            return False
