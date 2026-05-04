"""
Autenticação Microsoft Graph via device code flow.
Na primeira execução, abre uma URL — o usuário digita um código e faz login.
Depois disso o token é renovado automaticamente (cache local).
"""
import sys

import msal

from config import CLIENT_ID, AUTHORITY, SCOPES, TOKEN_CACHE_PATH

_app = None


def _get_app() -> msal.PublicClientApplication:
    global _app
    if _app:
        return _app

    if not CLIENT_ID:
        print("❌ CLIENT_ID não configurado em config.py")
        print("   Crie um App Registration no Azure Portal e cole o ID.")
        sys.exit(1)

    cache = msal.SerializableTokenCache()
    if TOKEN_CACHE_PATH.exists():
        cache.deserialize(TOKEN_CACHE_PATH.read_text())

    _app = msal.PublicClientApplication(
        CLIENT_ID,
        authority=AUTHORITY,
        token_cache=cache,
    )
    return _app


def _save_cache():
    app = _get_app()
    if app.token_cache.has_state_changed:
        TOKEN_CACHE_PATH.write_text(app.token_cache.serialize())



def get_token() -> str:
    """
    Retorna um access token válido.
    Usa cache se disponível, senão inicia device code flow.
    """
    app = _get_app()

    # Tenta token silencioso (cache/refresh)
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])
        if result and "access_token" in result:
            _save_cache()
            return result["access_token"]

    # Device code flow — usuário abre browser e digita código
    flow = app.initiate_device_flow(SCOPES)
    if "user_code" not in flow:
        raise RuntimeError(f"Erro de autenticação: {flow}")

    # Para uso em CLI legacy
    print()
    print("=" * 55)
    print("  AUTENTICAÇÃO MICROSOFT")
    print(f"  1. Abra: {flow['verification_uri']}")
    print(f"  2. Digite o código: {flow['user_code']}")
    print("  3. Faça login com sua conta Microsoft")
    print("=" * 55)
    print()

    result = app.acquire_token_by_device_flow(flow)

    if "access_token" not in result:
        err = result.get("error_description", str(result))
        raise RuntimeError(f"Login falhou: {err}")

    _save_cache()
    print("✅ Login realizado com sucesso!\n")
    return result["access_token"]


def get_device_code_flow():
    """
    Inicia o device code flow e retorna dict com 'verification_uri' e 'user_code'.
    NÃO executa o fluxo, apenas retorna os dados para uso em GUI.
    """
    app = _get_app()
    flow = app.initiate_device_flow(SCOPES)
    if "user_code" not in flow:
        raise RuntimeError(f"Erro de autenticação: {flow}")
    return flow, app

def complete_device_code_flow(flow, app):
    """
    Executa o fluxo de autenticação a partir de um device flow já iniciado.
    Retorna o access_token ou lança erro.
    """
    result = app.acquire_token_by_device_flow(flow)
    if "access_token" not in result:
        err = result.get("error_description", str(result))
        raise RuntimeError(f"Login falhou: {err}")
    _save_cache()
    return result["access_token"]
