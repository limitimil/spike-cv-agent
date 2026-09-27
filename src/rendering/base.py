from abc import ABC, abstractmethod
from pathlib import Path


class RendererError(RuntimeError):
    def __init__(self, code, message):
        self.code = code
        self.message = message
        super().__init__(code + ': ' + message)


class DocxToPdfRenderer(ABC):
    """Adapters write a complete PDF to a new output path or raise RendererError.

    Input documents and existing outputs must never be modified. Implementations
    own environment checks, timeouts, cleanup, and renderer-specific behavior.
    Metadata is JSON serializable and contains no live renderer objects.
    """
    @abstractmethod
    def validate_environment(self) -> None:
        pass

    @abstractmethod
    def metadata(self) -> dict:
        pass

    @abstractmethod
    def render(self, input_docx: Path, output_pdf: Path) -> None:
        pass
