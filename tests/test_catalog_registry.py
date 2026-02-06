import pytest

from src.catalog import _PROCESSOR_REGISTRY, get_processor, register_processor


def test_register_processor_adds_handler_and_get_processor_returns_it():
    processor_id = "unit_test_processor"

    @register_processor(processor_id)
    def _handler(item, output_dir, params):
        return (item, output_dir, params)

    try:
        assert get_processor(processor_id) is _handler
    finally:
        _PROCESSOR_REGISTRY.pop(processor_id, None)


def test_get_processor_raises_for_unknown_processor():
    with pytest.raises(KeyError):
        get_processor("unknown_processor_for_test")
