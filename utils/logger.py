"""
Logger module - Logging configuration for VCM
"""

import logging
import sys
from pathlib import Path


class VCMLogger:
    """Configure and manage logging for VCM"""
    
    _logger = None
    
    @staticmethod
    def setup(verbose: bool = False, log_file: Path = None) -> logging.Logger:
        """
        Setup and configure logger.
        
        Args:
            verbose: If True, set to DEBUG level; else INFO
            log_file: Optional file to write logs to
            
        Returns:
            Configured logger instance
        """
        if VCMLogger._logger is not None:
            return VCMLogger._logger
        
        # Create logger
        logger = logging.getLogger('vcm')
        
        # Set level
        level = logging.DEBUG if verbose else logging.INFO
        logger.setLevel(level)
        
        # Create formatter
        formatter = logging.Formatter(
            '%(asctime)s - %(name)s - %(levelname)s - %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )
        
        # Console handler
        console_handler = logging.StreamHandler(sys.stdout)
        console_handler.setLevel(level)
        console_handler.setFormatter(formatter)
        logger.addHandler(console_handler)
        
        # File handler (optional)
        if log_file:
            file_handler = logging.FileHandler(log_file)
            file_handler.setLevel(level)
            file_handler.setFormatter(formatter)
            logger.addHandler(file_handler)
        
        VCMLogger._logger = logger
        return logger
    
    @staticmethod
    def get_logger() -> logging.Logger:
        """
        Get the configured logger.
        
        Returns:
            Logger instance
        """
        if VCMLogger._logger is None:
            VCMLogger.setup()
        return VCMLogger._logger
    
    @staticmethod
    def debug(msg: str, *args, **kwargs):
        """Log debug message"""
        VCMLogger.get_logger().debug(msg, *args, **kwargs)
    
    @staticmethod
    def info(msg: str, *args, **kwargs):
        """Log info message"""
        VCMLogger.get_logger().info(msg, *args, **kwargs)
    
    @staticmethod
    def warning(msg: str, *args, **kwargs):
        """Log warning message"""
        VCMLogger.get_logger().warning(msg, *args, **kwargs)
    
    @staticmethod
    def error(msg: str, *args, **kwargs):
        """Log error message"""
        VCMLogger.get_logger().error(msg, *args, **kwargs)


if __name__ == "__main__":
    # Test logging
    logger = VCMLogger.setup(verbose=True)
    
    VCMLogger.debug("This is a debug message")
    VCMLogger.info("This is an info message")
    VCMLogger.warning("This is a warning message")
    VCMLogger.error("This is an error message")
