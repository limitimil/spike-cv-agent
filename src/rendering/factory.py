import platform

from .base import DocxToPdfRenderer, RendererError


class RendererFactory:
    @staticmethod
    def create() -> DocxToPdfRenderer:
        current = platform.system()
        if current != 'Windows':
            raise RendererError('UNSUPPORTED_PLATFORM',
                                current + ' rendering is unsupported. Use Windows with Microsoft Word installed.')
        # Lazy import keeps platform-specific dependencies behind this boundary.
        from .word_windows import WordAutomationRenderer
        return WordAutomationRenderer()
