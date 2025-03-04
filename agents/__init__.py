# agents/__init__.py
from .data_loader import DataLoaderAgent
from .content_generator import ContentGeneratorAgent
from .slide_builder import SlideBuilderAgent
from .plot_generator import PlotGeneratorAgent
from .report_assembler import ReportAssemblerAgent
from .ui_handler import UIHandlerAgent

__all__ = [
    'DataLoaderAgent',
    'ContentGeneratorAgent',
    'SlideBuilderAgent',
    'PlotGeneratorAgent',
    'ReportAssemblerAgent',
    'UIHandlerAgent'
]