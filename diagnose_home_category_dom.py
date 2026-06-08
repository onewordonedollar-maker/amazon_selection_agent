from __future__ import annotations

import json
import sys
import time
from pathlib import Path

from src.chrome_cdp import CDPClient, close_tab, open_tab


ROOT_URL = "https://www.amazon.com/gp/bestsellers/home-garden/ref=zg_bs_nav_home-garden_0"
OUTPUT = Path(__file__).resolve().parent / "outputs" / "home_category_dom_diagnostic.json"


def main():
    target_url = sys.argv[1] if len(sys.argv) > 1 else ROOT_URL
    target = open_tab(target_url)
    websocket_url = target.get("webSocketDebuggerUrl")
    if not websocket_url:
        raise RuntimeError("专用 Chrome 没有返回调试连接。")

    client = CDPClient(websocket_url, timeout=20)
    try:
        client.command("Runtime.enable")
        client.command("Page.enable")
        time.sleep(6)
        result = client.evaluate(
            r"""
            (() => {
              const browseGroups = Array.from(
                document.querySelectorAll('ul[class*="zg-browse-group"]')
              );
              const root =
                document.querySelector('#zg_browseRoot') ||
                document.querySelector('[data-csa-c-slot-id*="browse"]') ||
                document.querySelector('[class*="browseRoot"]') ||
                (browseGroups[0] && browseGroups[0].parentElement);
              const anchors = root ? Array.from(root.querySelectorAll('a[href]')) : [];
              const allAnchors = Array.from(document.querySelectorAll('a[href]'));
              const categoryAnchors = allAnchors.filter((anchor) => {
                try {
                  const url = new URL(anchor.href, location.href);
                  return (
                    /\/(?:zgbs|gp\/(?:bestsellers|new-releases))\//.test(url.pathname) &&
                    (/\/\d{5,}(?:\/|$)/.test(url.pathname) || url.searchParams.get('node'))
                  );
                } catch {
                  return false;
                }
              });
              const landmarks = Array.from(document.querySelectorAll('[id], [class]'))
                .filter((element) => {
                  const signature = `${element.id || ''} ${element.className || ''}`.toLowerCase();
                  return /browse|depart|category|zg|ranking/.test(signature);
                })
                .slice(0, 160)
                .map((element) => ({
                  tag: element.tagName,
                  id: element.id || '',
                  className: typeof element.className === 'string' ? element.className : '',
                  text: (element.innerText || element.textContent || '').trim().slice(0, 240)
                }));
              return {
                pageTitle: document.title,
                pageUrl: location.href,
                rootFound: Boolean(root),
                rootTag: root ? root.tagName : '',
                rootClass: root ? root.className : '',
                browseGroupCount: browseGroups.length,
                browseGroups: browseGroups.slice(0, 20).map((group) => ({
                  className: group.className || '',
                  childCount: group.children ? group.children.length : 0,
                  text: (group.innerText || group.textContent || '').trim().slice(0, 1200)
                })),
                selected: root
                  ? Array.from(root.querySelectorAll('.zg_selected, [aria-current="page"]'))
                      .slice(0, 10)
                      .map((element) => ({
                        tag: element.tagName,
                        text: (element.innerText || element.textContent || '').trim(),
                        className: element.className || '',
                        parentTag: element.parentElement ? element.parentElement.tagName : '',
                        parentClass: element.parentElement ? element.parentElement.className || '' : ''
                      }))
                  : [],
                bodyTextStart: (document.body.innerText || '').trim().slice(0, 2500),
                categoryAnchors: categoryAnchors.slice(0, 200).map((anchor) => {
                  const rect = anchor.getBoundingClientRect();
                  const parents = [];
                  let parent = anchor.parentElement;
                  for (let depth = 0; parent && depth < 5; depth += 1) {
                    parents.push({
                      tag: parent.tagName,
                      id: parent.id || '',
                      className: typeof parent.className === 'string' ? parent.className : ''
                    });
                    parent = parent.parentElement;
                  }
                  return {
                    text: (anchor.innerText || anchor.textContent || '').trim(),
                    href: anchor.href,
                    x: Math.round(rect.left),
                    y: Math.round(rect.top),
                    visible: rect.width > 0 && rect.height > 0,
                    ariaCurrent: anchor.getAttribute('aria-current') || '',
                    parents
                  };
                }),
                landmarks,
                anchors: anchors.slice(0, 120).map((anchor) => {
                  const rect = anchor.getBoundingClientRect();
                  const li = anchor.closest('li');
                  return {
                    text: (anchor.innerText || anchor.textContent || '').trim(),
                    href: anchor.href,
                    x: Math.round(rect.left),
                    y: Math.round(rect.top),
                    parentTag: anchor.parentElement ? anchor.parentElement.tagName : '',
                    parentClass: anchor.parentElement ? anchor.parentElement.className || '' : '',
                    liClass: li ? li.className || '' : '',
                    ariaCurrent: anchor.getAttribute('aria-current') || ''
                  };
                })
              };
            })()
            """,
            timeout=20,
        )
        OUTPUT.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"诊断完成：{OUTPUT}")
        print(f"找到类目区域：{result.get('rootFound')}，链接数：{len(result.get('anchors', []))}")
    finally:
        client.close()
        close_tab(target.get("id"))


if __name__ == "__main__":
    main()
