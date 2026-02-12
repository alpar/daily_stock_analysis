# feishu_doc.py
# -*- coding: utf-8 -*-
import logging
import json
import sys
import lark_oapi as lark
from lark_oapi.api.docx.v1 import *
from typing import List, Dict, Any, Optional
from src.config import get_config

try:
    from markdown_it import MarkdownIt
    from markdown_it.token import Token
except Exception:
    MarkdownIt = None
    Token = None

# friendly hint when optional parser is missing
MD_PARSER_AVAILABLE = MarkdownIt is not None
if not MD_PARSER_AVAILABLE:
    msg = (
        "Optional dependency 'markdown-it-py' not found. "
        "Install it to enable richer Markdown parsing:\n"
        "  pip install markdown-it-py"
    )
    logger = logging.getLogger(__name__)
    logger.warning(msg)
    try:
        print(msg, file=sys.stderr)
    except Exception:
        pass

logger = logging.getLogger(__name__)


class FeishuDocManager:
    """飞书云文档管理器 (基于官方 SDK lark-oapi)"""

    def __init__(self):
        self.config = get_config()
        self.app_id = self.config.feishu_app_id
        self.app_secret = self.config.feishu_app_secret
        self.folder_token = self.config.feishu_folder_token

        # 初始化 SDK 客户端
        # SDK 会自动处理 tenant_access_token 的获取和刷新，无需人工干预
        if self.is_configured():
            self.client = lark.Client.builder() \
                .app_id(self.app_id) \
                .app_secret(self.app_secret) \
                .log_level(lark.LogLevel.INFO) \
                .build()
        else:
            self.client = None

    def is_configured(self) -> bool:
        """检查配置是否完整"""
        return bool(self.app_id and self.app_secret and self.folder_token)

    def create_daily_doc(self, title: str, content_md: str) -> Optional[str]:
        """
        创建日报文档
        """
        if not self.client or not self.is_configured():
            logger.warning("飞书 SDK 未初始化或配置缺失，跳过创建")
            return None

        try:
            # 1. 创建文档
            # 使用官方 SDK 的 Builder 模式构造请求
            create_request = CreateDocumentRequest.builder() \
                .request_body(CreateDocumentRequestBody.builder()
                              .folder_token(self.folder_token)
                              .title(title)
                              .build()) \
                .build()

            response = self.client.docx.v1.document.create(create_request)

            if not response.success():
                logger.error(f"创建文档失败: {response.code} - {response.msg} - {response.error}")
                return None

            doc_id = response.data.document.document_id
            # 这里的 domain 只是为了生成链接，实际访问会重定向
            doc_url = f"https://feishu.cn/docx/{doc_id}"
            logger.info(f"飞书文档创建成功: {title} (ID: {doc_id})")

            # 2. 解析 Markdown 并写入内容
            # 将 Markdown 转换为 SDK 需要的 Block 对象列表
            blocks = self._markdown_to_sdk_blocks(content_md)

            # 飞书 API 限制每次写入 Block 数量（建议 50 个左右），分批写入
            batch_size = 50
            doc_block_id = doc_id  # 文档本身也是一个 block

            for i in range(0, len(blocks), batch_size):
                batch_blocks = blocks[i:i + batch_size]

                # 构造批量添加块的请求
                batch_add_request = CreateDocumentBlockChildrenRequest.builder() \
                    .document_id(doc_id) \
                    .block_id(doc_block_id) \
                    .request_body(CreateDocumentBlockChildrenRequestBody.builder()
                                  .children(batch_blocks)  # SDK 需要 Block 对象列表
                                  .index(-1)  # 追加到末尾
                                  .build()) \
                    .build()

                write_resp = self.client.docx.v1.document_block_children.create(batch_add_request)

                if not write_resp.success():
                    logger.error(f"写入文档内容失败(批次{i}): {write_resp.code} - {write_resp.msg}")

            logger.info(f"文档内容写入完成")
            return doc_url

        except Exception as e:
            logger.error(f"飞书文档操作异常: {e}")
            import traceback
            logger.error(traceback.format_exc())
            return None

    def _markdown_to_sdk_blocks(self, md_text: str) -> List[Block]:
        """
        将简单的 Markdown 转换为飞书 SDK 的 Block 对象
        """
        blocks: List[Block] = []

        def build_text_obj(content: str) -> Text:
            text_run = TextRun.builder() \
                .content(content) \
                .text_element_style(TextElementStyle.builder().build()) \
                .build()

            text_element = TextElement.builder() \
                .text_run(text_run) \
                .build()

            return Text.builder() \
                .elements([text_element]) \
                .style(TextStyle.builder().build()) \
                .build()

        def extract_inline_text(inline_token: Token) -> str:
            if not getattr(inline_token, 'children', None):
                return inline_token.content or ''

            parts: List[str] = []
            for child in inline_token.children:
                t = child.type
                if t == 'text':
                    parts.append(child.content)
                elif t == 'code_inline':
                    parts.append('`' + child.content + '`')
                elif t == 'image':
                    src = child.attrGet('src') or ''
                    alt = child.attrGet('alt') or 'image'
                    parts.append(f'[图片: {alt}]({src})')
                elif t == 'link_open':
                    # mark start of link
                    parts.append('[')
                elif t == 'link_close':
                    parts.append(']')
                elif t in ('strong_open', 'strong_close'):
                    parts.append('**')
                elif t in ('em_open', 'em_close'):
                    parts.append('*')
                else:
                    if hasattr(child, 'content') and child.content:
                        parts.append(child.content)

            return ''.join(parts)

        if MarkdownIt is None:
            # fallback to original simple line parser
            lines = md_text.split('\n')
            for line in lines:
                s = line.strip()
                if not s:
                    continue
                if s.startswith('# '):
                    text_obj = build_text_obj(s[2:])
                    blocks.append(Block.builder().block_type(3).heading1(text_obj).build())
                    continue
                if s.startswith('## '):
                    text_obj = build_text_obj(s[3:])
                    blocks.append(Block.builder().block_type(4).heading2(text_obj).build())
                    continue
                if s.startswith('### '):
                    text_obj = build_text_obj(s[4:])
                    blocks.append(Block.builder().block_type(5).heading3(text_obj).build())
                    continue
                if s.startswith('---'):
                    blocks.append(Block.builder().block_type(22).divider(Divider.builder().build()).build())
                    continue
                text_obj = build_text_obj(s)
                blocks.append(Block.builder().block_type(2).text(text_obj).build())

            return blocks

        md = MarkdownIt().enable('table')
        tokens = md.parse(md_text)

        i = 0
        while i < len(tokens):
            token = tokens[i]
            ttype = token.type

            if ttype == 'heading_open':
                level = int(token.tag[1]) if token.tag and len(token.tag) > 1 else 1
                inline = tokens[i + 1]
                content = extract_inline_text(inline)
                text_obj = build_text_obj(content)
                if level == 1:
                    blocks.append(Block.builder().block_type(3).heading1(text_obj).build())
                elif level == 2:
                    blocks.append(Block.builder().block_type(4).heading2(text_obj).build())
                else:
                    blocks.append(Block.builder().block_type(5).heading3(text_obj).build())
                i += 3
                continue

            if ttype == 'paragraph_open':
                inline = tokens[i + 1]
                content = extract_inline_text(inline)
                text_obj = build_text_obj(content)
                blocks.append(Block.builder().block_type(2).text(text_obj).build())
                i += 3
                continue

            if ttype == 'fence':
                code = token.content.rstrip('\n')
                # 将代码块作为预格式化文本写入
                text_obj = build_text_obj('```\n' + code + '\n```')
                blocks.append(Block.builder().block_type(2).text(text_obj).build())
                i += 1
                continue

            if ttype in ('bullet_list_open', 'ordered_list_open'):
                ordered = (ttype == 'ordered_list_open')
                j = i + 1
                index_counter = 1
                while j < len(tokens) and tokens[j].type not in ('bullet_list_close', 'ordered_list_close'):
                    if tokens[j].type == 'list_item_open':
                        inline_token = None
                        for k in range(j, min(j + 6, len(tokens))):
                            if tokens[k].type == 'inline':
                                inline_token = tokens[k]
                                break
                        if inline_token:
                            content = extract_inline_text(inline_token)
                            prefix = f"{index_counter}. " if ordered else "- "
                            text_obj = build_text_obj(prefix + content)
                            blocks.append(Block.builder().block_type(2).text(text_obj).build())
                        index_counter += 1
                    j += 1
                i = j + 1
                continue

            if ttype == 'blockquote_open':
                j = i + 1
                lines = []
                while j < len(tokens) and tokens[j].type != 'blockquote_close':
                    if tokens[j].type == 'paragraph_open':
                        inline = tokens[j + 1]
                        lines.append('> ' + extract_inline_text(inline))
                        j += 3
                        continue
                    j += 1
                text_obj = build_text_obj('\n'.join(lines))
                blocks.append(Block.builder().block_type(2).text(text_obj).build())
                i = j + 1
                continue

            if ttype == 'table_open':
                j = i + 1
                headers = []
                rows: List[List[str]] = []
                current_row: List[str] = []
                while j < len(tokens) and tokens[j].type != 'table_close':
                    if tokens[j].type == 'th_open':
                        inline = tokens[j + 1]
                        headers.append(extract_inline_text(inline))
                        j += 3
                        continue
                    if tokens[j].type == 'td_open':
                        inline = tokens[j + 1]
                        current_row.append(extract_inline_text(inline))
                        j += 3
                        # if next is tr_close, push row
                        if j < len(tokens) and tokens[j].type == 'tr_close':
                            if current_row:
                                rows.append(current_row)
                            current_row = []
                        continue
                    j += 1

                table_lines: List[str] = []
                if headers:
                    table_lines.append('| ' + ' | '.join(headers) + ' |')
                    table_lines.append('|' + '---|' * len(headers))
                for r in rows:
                    table_lines.append('| ' + ' | '.join(r) + ' |')

                if table_lines:
                    text_obj = build_text_obj('\n'.join(table_lines))
                    blocks.append(Block.builder().block_type(2).text(text_obj).build())

                i = j + 1
                continue

            i += 1

        return blocks
