"""Azure OpenAI is the only LLM provider; without it everything runs in mock mode."""
from __future__ import annotations

from workspace import config, extractor_adapter


def test_mock_without_azure_and_no_groq_fallback(monkeypatch):
    monkeypatch.delenv("WORKSPACE_LLM", raising=False)
    monkeypatch.setenv("GROQ_API_KEY", "x")
    assert config.llm_mode() == "mock" and config.chat_models() == [] and config.active_model() is None


def test_azure_from_api_base_style_variables(monkeypatch):
    monkeypatch.delenv("WORKSPACE_LLM", raising=False)
    monkeypatch.setenv("OPENAI_API_KEY", "k")
    monkeypatch.setenv("API_BASE", "https://res-1.openai.azure.com/openai/deployments/gpt-5")
    monkeypatch.setenv("LANGUAGE_MODEL", "openai/gpt-5")
    assert config.llm_mode() == "azure" and config.azure_endpoint() == "https://res-1.openai.azure.com"
    assert config.azure_deployment() == "gpt-5" and config.active_model() == "azure:gpt-5"
    [(label, llm)] = config.chat_models(max_tokens=100)
    assert label == "azure:gpt-5" and llm.deployment_name == "gpt-5" and llm.max_tokens == 16000


def test_legacy_extractor_decider_is_pointed_at_azure(monkeypatch):
    from xlsx_template import decider as dec
    monkeypatch.delenv("GROQ_API_KEY", raising=False)
    monkeypatch.setenv("AZURE_OPENAI_API_KEY", "k")
    monkeypatch.setenv("AZURE_OPENAI_ENDPOINT", "https://res-1.openai.azure.com")
    orig = dec.LLMDecider
    with extractor_adapter._azure_decider():
        d = dec.LLMDecider(model="openai/gpt-oss-120b")
        llm = d._llm(d.model, 500)
        assert d.model == config.azure_deployment() and d.fallback_model is None
        assert type(llm).__name__ == "AzureChatOpenAI" and llm.model_kwargs["response_format"]["type"] == "json_object"
    assert dec.LLMDecider is orig
