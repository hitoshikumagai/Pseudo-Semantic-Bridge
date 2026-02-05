import os
from pathlib import Path
from typing import Optional

from src.catalog import register_processor


def _resolve_api_key(token_env: str) -> Optional[str]:
    if token_env:
        return os.environ.get(token_env)
    # Default to OpenAI recommended env var
    return os.environ.get("OPENAI_API_KEY")


@register_processor("agent_external_api")
def external_api_agent(item, output_dir, params):
    """
    Text-generation agent using the OpenAI SDK.
    params example:
      {
        "model": "gpt-4.1",
        "prompt": "Summarize: {item_name}",
        "token_env": "OPENAI_API_KEY",
        "save_output": true
      }
    """
    try:
        from openai import OpenAI
    except Exception as e:
        print(f"      ❌ OpenAI SDK import failed: {e}")
        return

    model = str(params.get("model", "gpt-4.1"))
    prompt_template = str(params.get("prompt", "Summarize: {item_name}"))
    token_env = str(params.get("token_env", "OPENAI_API_KEY")).strip()
    save_output = bool(params.get("save_output", True))

    api_key = _resolve_api_key(token_env)
    if not api_key:
        print("      ⚠️ API key not found. Check your environment variables.")
        return

    client = OpenAI(api_key=api_key)

    prompt = prompt_template.format(
        item_name=item.name,
        item_extension=item.extension,
    )

    try:
        response = client.responses.create(
            model=model,
            input=prompt,
        )
        output_text = getattr(response, "output_text", None)
        if not output_text:
            output_text = str(response)

        print("      ✅ OpenAI response received.")

        if save_output:
            out_dir = Path(output_dir)
            out_dir.mkdir(parents=True, exist_ok=True)
            out_path = out_dir / f"{item.name}.openai.txt"
            out_path.write_text(output_text, encoding="utf-8")
            print(f"      📝 Saved: {out_path}")
    except Exception as e:
        print(f"      ❌ OpenAI call failed: {e}")
