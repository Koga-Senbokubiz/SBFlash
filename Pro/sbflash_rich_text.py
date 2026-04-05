# -*- coding: utf-8 -*-
"""
SBFlash用 簡易リッチテキスト表示モジュール
対応タグ:
    <b> ... </b>       : 太字
    <red> ... </red>   : 赤文字
    <big> ... </big>   : 拡大文字
"""
from __future__ import annotations
import tkinter as tk
from tkinter import font
from typing import List, Tuple, Set

class SimpleRichTextParser:
    SUPPORTED_TAGS = {"b", "red", "big"}

    def parse(self, content: str) -> List[Tuple[str, Set[str]]]:
        if content is None:
            return []
        s = str(content)
        result: List[Tuple[str, Set[str]]] = []
        style_stack: List[str] = []
        buffer: List[str] = []
        i = 0
        n = len(s)
        while i < n:
            if s[i] == "<":
                close_pos = s.find(">", i)
                if close_pos != -1:
                    tag_token = s[i + 1:close_pos].strip()
                    if self._is_supported_tag_token(tag_token):
                        if buffer:
                            result.append(("".join(buffer), set(style_stack)))
                            buffer = []
                        if tag_token.startswith("/"):
                            tag_name = tag_token[1:].strip()
                            self._pop_style(style_stack, tag_name)
                        else:
                            style_stack.append(tag_token)
                        i = close_pos + 1
                        continue
            buffer.append(s[i])
            i += 1
        if buffer:
            result.append(("".join(buffer), set(style_stack)))
        return self._merge_adjacent(result)

    def _is_supported_tag_token(self, token: str) -> bool:
        if not token:
            return False
        if token.startswith("/"):
            return token[1:] in self.SUPPORTED_TAGS
        return token in self.SUPPORTED_TAGS

    def _pop_style(self, style_stack: List[str], tag_name: str) -> None:
        for idx in range(len(style_stack) - 1, -1, -1):
            if style_stack[idx] == tag_name:
                del style_stack[idx]
                return

    def _merge_adjacent(self, items: List[Tuple[str, Set[str]]]) -> List[Tuple[str, Set[str]]]:
        if not items:
            return items
        merged: List[Tuple[str, Set[str]]] = []
        prev_text, prev_styles = items[0]
        for text, styles in items[1:]:
            if styles == prev_styles:
                prev_text += text
            else:
                merged.append((prev_text, prev_styles))
                prev_text, prev_styles = text, styles
        merged.append((prev_text, prev_styles))
        return merged

def apply_rich_text_to_text_widget(
    text_widget: tk.Text,
    content: str,
    base_font_family: str = "Yu Gothic UI",
    base_font_size: int = 12,
    big_font_size: int = 16
) -> None:
    parser = SimpleRichTextParser()
    base_font_obj = font.Font(family=base_font_family, size=base_font_size)
    bold_font = font.Font(family=base_font_family, size=base_font_size, weight="bold")
    big_font_obj = font.Font(family=base_font_family, size=big_font_size)
    big_bold_font = font.Font(family=base_font_family, size=big_font_size, weight="bold")

    text_widget.tag_configure("normal", font=base_font_obj, foreground="black")
    text_widget.tag_configure("bold", font=bold_font, foreground="black")
    text_widget.tag_configure("red", font=base_font_obj, foreground="red")
    text_widget.tag_configure("big", font=big_font_obj, foreground="black")
    text_widget.tag_configure("bold_red", font=bold_font, foreground="red")
    text_widget.tag_configure("big_red", font=big_font_obj, foreground="red")
    text_widget.tag_configure("big_bold", font=big_bold_font, foreground="black")
    text_widget.tag_configure("big_bold_red", font=big_bold_font, foreground="red")

    current_state = str(text_widget.cget("state"))
    if current_state == "disabled":
        text_widget.configure(state="normal")
    text_widget.delete("1.0", "end")

    for text_part, style_set in parser.parse(content):
        has_bold = "b" in style_set
        has_red = "red" in style_set
        has_big = "big" in style_set
        if has_big and has_bold and has_red:
            tag_name = "big_bold_red"
        elif has_big and has_bold:
            tag_name = "big_bold"
        elif has_big and has_red:
            tag_name = "big_red"
        elif has_bold and has_red:
            tag_name = "bold_red"
        elif has_big:
            tag_name = "big"
        elif has_bold:
            tag_name = "bold"
        elif has_red:
            tag_name = "red"
        else:
            tag_name = "normal"
        text_widget.insert("end", text_part, tag_name)

    text_widget.see("1.0")
    if current_state == "disabled":
        text_widget.configure(state="disabled")
