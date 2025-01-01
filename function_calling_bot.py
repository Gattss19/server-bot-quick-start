"""

Sample bot that demonstrates how to use OpenAI function calling with the Poe API.

"""

from __future__ import annotations

import json
from typing import AsyncIterable

import fastapi_poe as fp
from modal import App, Image, asgi_app


def get_current_weather(location, unit="fahrenheit"):
    """Get the current weather in a given location"""
    if "tokyo" in location.lower():
        return json.dumps({"location": "Tokyo", "temperature": "11", "unit": unit})
    elif "san francisco" in location.lower():
        return json.dumps(
            {"location": "San Francisco", "temperature": "72", "unit": unit}
        )
    elif "paris" in location.lower():
        return json.dumps({"location": "Paris", "temperature": "22", "unit": unit})
    else:
        return json.dumps({"location": location, "temperature": "unknown"})


tools_executables = [get_current_weather]

tools_dict_list = [
    {
        "type": "function",
        "function": {
            "name": "get_current_weather",
            "description": "Get the current weather in a given location",
            "parameters": {
                "type": "object",
                "properties": {
                    "location": {
                        "type": "string",
                        "description": "The city and state, e.g. San Francisco, CA",
                    },
                    "unit": {"type": "string", "enum": ["celsius", "fahrenheit"]},
                },
                "required": ["location"],
            },
        },
    }
]
tools = [fp.ToolDefinition(**tools_dict) for tools_dict in tools_dict_list]


class GPT35FunctionCallingBot(fp.PoeBot):
    async def get_response(
        self, request: fp.QueryRequest
    ) -> AsyncIterable[fp.PartialResponse]:
        async for msg in fp.stream_request(
            request,
            "GPT-3.5-Turbo",
            request.access_key,
            tools=tools,
            tool_executables=tools_executables,
        ):
            yield msg

    async def get_settings(self, setting: fp.SettingsRequest) -> fp.SettingsResponse:
        return fp.SettingsResponse(server_bot_dependencies={"GPT-3.5-Turbo": 2})
    async def get_settings(self, setting: fp.SettingsRequest) -> fp.SettingsResponse:
        return fp.SettingsResponse(
        introduction_message="Привет! Я ваш новый бот-ассистент. Задайте мне любой вопрос!"
    )



REQUIREMENTS = ["fastapi-poe==0.0.48"]
image = Image.debian_slim().pip_install(*REQUIREMENTS)
app = App("function-calling-poe")


@app.function(image=image)
@asgi_app()
def fastapi_app():
    bot = GPT35FunctionCallingBot()
    # see https://creator.poe.com/docs/quick-start#configuring-the-access-credentials
    # app = fp.make_app(bot, access_key=<6dXnPogSx0g73Yvm3xLWENaG44mgVIJq>, bot_name=<AuditAssistantKrs>)
    app = fp.make_app(bot, allow_without_key=True)
    return app
