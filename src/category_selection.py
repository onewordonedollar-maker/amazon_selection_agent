from collections.abc import Callable


def category_path_parts(path: str) -> list[str]:
    return [part.strip() for part in path.split(">") if part.strip()]


def path_covers(parent: str, child: str) -> bool:
    return child == parent or child.startswith(f"{parent} > ")


def category_path_selected(path: str, selected_paths: list[str]) -> bool:
    return any(path_covers(selected_path, path) for selected_path in selected_paths)


def get_category_node(tree: dict, path: str) -> dict:
    current = tree
    node = {}
    for part in category_path_parts(path):
        node = current.get(part, {}) if isinstance(current, dict) else {}
        current = node.get("children", {}) if isinstance(node, dict) else {}
    return node if isinstance(node, dict) else {}


def iter_child_paths(tree: dict, path: str):
    node = get_category_node(tree, path)
    for child_name, child_node in node.get("children", {}).items():
        yield f"{path} > {child_name}", child_node


def compact_category_selection(paths: list[str]) -> list[str]:
    unique_paths = list(dict.fromkeys(path for path in paths if path))
    order_by_path = {path: index for index, path in enumerate(unique_paths)}
    ordered_paths = sorted(unique_paths, key=lambda path: (path.count(" > "), order_by_path[path]))
    compact: list[str] = []
    for path in ordered_paths:
        if any(path_covers(parent, path) for parent in compact):
            continue
        compact.append(path)
    return compact


def subtract_category_branch(tree: dict, selected_branch: str, removed_path: str) -> list[str]:
    if not path_covers(selected_branch, removed_path):
        return [selected_branch]
    if selected_branch == removed_path:
        return []

    kept: list[str] = []
    cursor = selected_branch
    while cursor != removed_path:
        next_remove = None
        for child_path, _child_node in iter_child_paths(tree, cursor):
            if path_covers(child_path, removed_path):
                next_remove = child_path
            else:
                kept.append(child_path)
        if not next_remove:
            break
        cursor = next_remove
    return kept


def toggle_compact_category_selection(
    tree: dict,
    selected_paths: list[str],
    path: str,
    selected: bool,
) -> list[str]:
    compact = compact_category_selection(selected_paths)
    if selected:
        if category_path_selected(path, compact):
            return compact
        return compact_category_selection([*compact, path])

    updated: list[str] = []
    for selected_path in compact:
        if path_covers(path, selected_path):
            continue
        if path_covers(selected_path, path):
            updated.extend(subtract_category_branch(tree, selected_path, path))
            continue
        updated.append(selected_path)
    return compact_category_selection(updated)


def selected_compact_category_paths(
    tree: dict,
    is_selected: Callable[[str], bool],
    prefix: str = "",
    seen: set[int] | None = None,
) -> list[str]:
    if not isinstance(tree, dict) or not tree:
        return []
    seen = seen or set()
    tree_id = id(tree)
    if tree_id in seen:
        return []
    seen.add(tree_id)

    selected_paths: list[str] = []
    for name, node in tree.items():
        path = f"{prefix} > {name}" if prefix else name
        if is_selected(path):
            selected_paths.append(path)
            continue
        children = node.get("children", {}) if isinstance(node, dict) else {}
        selected_paths.extend(selected_compact_category_paths(children, is_selected, path, seen))
    return selected_paths
