from collections.abc import Generator
from typing import Any

from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage

import requests
import pandas as pd
from io import BytesIO

class ExcelToMdTool(Tool):
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage]:
        file_url = tool_parameters['file_url']
        file_content = requests.get(file_url).content
        file_stream = BytesIO(file_content)
        xl = pd.read_excel(file_stream, sheet_name=None)
        result = []
        print(1)
        for sheet_name, df in xl.items():
            md = df.to_markdown(index=False)
            result.append(f"### Sheet: {sheet_name}\n\n{md}")
        print(2)
        print(str(result))
        print(3)
        final_markdown = "\n\n---\n\n".join(result)
        yield self.create_text_message(final_markdown)
