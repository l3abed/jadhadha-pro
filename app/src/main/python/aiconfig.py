# -*- coding: utf-8 -*-
"""
aiconfig.py — وجهة نداء الذكاء الاصطناعي.
=========================================
إمّا اتصال مباشر بـGoogle Gemini بمفتاح المستخدم، أو عبر وسيط خلفي يحمل المفتاح
(لا يُشحن المفتاح في التطبيق). يُحدَّد الوسيط عبر متغيّر البيئة JADHADHA_AI_BASE
الذي يضبطه bridge.py من إعدادات التطبيق.
"""
from __future__ import annotations

import os

DEFAULT_BASE = "https://generativelanguage.googleapis.com"


def base_url() -> str:
    """جذر خدمة Gemini: المباشر افتراضاً، أو عنوان الوسيط إن كان مضبوطاً."""
    return (os.environ.get("JADHADHA_AI_BASE") or DEFAULT_BASE).rstrip("/")


def via_proxy() -> bool:
    return bool((os.environ.get("JADHADHA_AI_BASE") or "").strip())


def endpoint(model: str, method: str = "generateContent") -> str:
    return f"{base_url()}/v1beta/models/{model}:{method}"
