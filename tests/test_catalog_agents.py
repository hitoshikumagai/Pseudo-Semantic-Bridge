from src.catalog import get_processor


def test_get_processor_loads_agent():
    handler = get_processor("agent_external_api")
    assert callable(handler)
