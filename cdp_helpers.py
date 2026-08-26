"""
Helpers para interagir com elementos dentro de Shadow DOM FECHADO via CDP
(Chrome DevTools Protocol), que o Playwright normal (Locator/querySelector)
não consegue atravessar.

Cópia canônica deste arquivo — todo projeto novo em automacao_adm/ que
encontrar uma tela do Innovaro renderizada com poucos elementos no DOM (mas
com bastante conteúdo visual) deve copiar este arquivo para o próprio
projeto. Ver BOAS_PRATICAS_INNOVARO_PLAYWRIGHT.md, seção 24, para o contexto
completo da descoberta (feita originalmente em duplicacao_cadastro/).

Uso básico:

    cdp = context.new_cdp_session(page)
    cdp.send("DOM.enable")

    btn = find_node(cdp, lambda tag, a: tag == "BUTTON" and a.get("aria-label") == "Pesquisar")
    click_node(cdp, btn)
"""


def get_flattened_nodes(cdp):
    """Retorna a lista de todos os nós do documento, atravessando shadow DOM fechado."""
    doc = cdp.send("DOM.getFlattenedDocument", {"depth": -1, "pierce": True})
    return doc["nodes"]


def node_attrs(node):
    attrs = node.get("attributes", [])
    return dict(zip(attrs[0::2], attrs[1::2])) if attrs else {}


def find_node(cdp, predicate):
    """Retorna o primeiro nó (dict do CDP) cujos attrs satisfazem predicate(nodeName, attrs)."""
    for n in get_flattened_nodes(cdp):
        if predicate(n.get("nodeName", ""), node_attrs(n)):
            return n
    return None


def find_all_nodes(cdp, predicate):
    return [n for n in get_flattened_nodes(cdp) if predicate(n.get("nodeName", ""), node_attrs(n))]


def _resolve_object_id(cdp, node):
    resolved = cdp.send("DOM.resolveNode", {"nodeId": node["nodeId"]})
    return resolved["object"]["objectId"]


def click_node(cdp, node):
    """Clica no elemento via JS .click() (funciona mesmo dentro de shadow DOM fechado)."""
    object_id = _resolve_object_id(cdp, node)
    cdp.send("Runtime.callFunctionOn", {
        "objectId": object_id,
        "functionDeclaration": (
            "function() { this.scrollIntoView({block: 'center', inline: 'center'}); "
            "this.click(); }"
        ),
    })


def focus_node(cdp, node):
    object_id = _resolve_object_id(cdp, node)
    cdp.send("Runtime.callFunctionOn", {
        "objectId": object_id,
        "functionDeclaration": "function() { this.focus(); }",
    })


def read_value(cdp, node):
    object_id = _resolve_object_id(cdp, node)
    result = cdp.send("Runtime.callFunctionOn", {
        "objectId": object_id,
        "functionDeclaration": "function() { return this.value; }",
        "returnByValue": True,
    })
    return result["result"]["value"]


def search_text(cdp, query):
    """
    Busca nós (inclusive dentro de shadow DOM) por texto/seletor via
    DOM.performSearch — útil para achar itens de menu que só têm texto
    visível, sem aria-label/name (ex.: "Duplicar registro"). Retorna uma
    lista de nodeId (não nós completos).

    IMPORTANTE: precisa "aquecer" a árvore do DOM (DOM.getFlattenedDocument)
    antes de DOM.performSearch nessa mesma sessão CDP — sem isso, os nodeId
    retornados dão erro "No node with given id found" em DOM.resolveNode
    depois. Por isso chamamos get_flattened_nodes aqui sempre.
    """
    get_flattened_nodes(cdp)
    result = cdp.send("DOM.performSearch", {"query": query, "includeUserAgentShadowDOM": True})
    search_id = result["searchId"]
    count = result["resultCount"]
    node_ids = []
    if count:
        node_ids = cdp.send("DOM.getSearchResults", {
            "searchId": search_id, "fromIndex": 0, "toIndex": count,
        })["nodeIds"]
    cdp.send("DOM.discardSearchResults", {"searchId": search_id})
    return node_ids


def click_node_id(cdp, node_id):
    """
    Clica em um elemento a partir do nodeId direto (ex.: vindo de
    search_text). Se o nó for um nó de texto (ou não for clicável em si),
    sobe até o <button> mais próximo antes de clicar — texto de menu/botão
    geralmente vem num nó de texto dentro de um <span>, não no <button>.
    """
    object_id = cdp.send("DOM.resolveNode", {"nodeId": node_id})["object"]["objectId"]
    cdp.send("Runtime.callFunctionOn", {
        "objectId": object_id,
        "functionDeclaration": (
            "function() { "
            "const start = this.nodeType === 3 ? this.parentElement : this; "
            "const el = start.closest('button') || start; "
            "el.scrollIntoView({block: 'center', inline: 'center'}); "
            "el.click(); }"
        ),
    })


def click_text(cdp, query):
    """
    Combina search_text + click_node_id: acha o primeiro nó cujo texto bate
    com `query` e clica no <button> mais próximo dele. Uso típico: itens de
    menu ou botões de toolbar identificados só pelo texto visível (sem
    aria-label/name), ex.: "Duplicar registro", "Visualizar".
    """
    node_ids = search_text(cdp, query)
    if not node_ids:
        raise RuntimeError(f"Nenhum elemento com texto '{query}' encontrado.")
    click_node_id(cdp, node_ids[0])


def dump_nodes_resumo(cdp, apenas_com_atributos: bool = True):
    """
    Utilitário de exploração: lista (nodeName, attrs) de todos os nós, para
    mapear rapidamente o que existe numa tela nova (ex.: rodar isso logo
    depois de abrir uma tela e dar print/json.dumps no resultado).
    """
    nodes = get_flattened_nodes(cdp)
    resumo = []
    for n in nodes:
        attrs = node_attrs(n)
        if apenas_com_atributos and not attrs:
            continue
        resumo.append({"nodeName": n.get("nodeName", ""), "attrs": attrs})
    return resumo
