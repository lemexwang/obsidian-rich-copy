const obsidian = require('obsidian');

class RichCopyPlugin extends obsidian.Plugin {
    async onload() {
        console.log('Loading Rich Text Copy Tool');

        this.addCommand({
            id: 'copy-as-rich-text',
            name: 'Copy current note as Rich Text (for Outlook/Teams)',
            checkCallback: (checking) => {
                const view = this.app.workspace.getActiveViewOfType(obsidian.MarkdownView);
                if (view) {
                    if (!checking) {
                        this.copyRichText(view);
                    }
                    return true;
                }
                return false;
            }
        });

        this.registerEvent(
            this.app.workspace.on('layout-change', () => {
                this.injectButtons();
            })
        );

        this.app.workspace.onLayoutReady(() => {
            this.injectButtons();
        });
    }

    injectButtons() {
        this.app.workspace.iterateAllLeaves((leaf) => {
            if (leaf.view instanceof obsidian.MarkdownView) {
                const container = leaf.view.containerEl;
                if (container.querySelector('.rich-copy-floating-btn')) return;

                const btn = container.createEl('button', {
                    cls: 'clickable-icon rich-copy-floating-btn',
                    attr: { 'aria-label': 'Copy as Rich Text' }
                });

                Object.assign(btn.style, {
                    position: 'absolute',
                    bottom: '20px',
                    right: '30px',
                    zIndex: '100',
                    backgroundColor: 'var(--interactive-accent)',
                    color: 'white',
                    borderRadius: '50%',
                    width: '40px',
                    height: '40px',
                    display: 'flex',
                    alignItems: 'center',
                    justifyContent: 'center',
                    boxShadow: '0 4px 10px rgba(0,0,0,0.3)',
                    border: 'none',
                    cursor: 'pointer',
                    transition: 'transform 0.1s ease'
                });

                obsidian.setIcon(btn, 'copy');

                btn.onclick = (e) => {
                    e.stopPropagation();
                    this.copyRichText(leaf.view);
                    
                    obsidian.setIcon(btn, 'check');
                    btn.style.backgroundColor = '#48bb78';
                    setTimeout(() => {
                        obsidian.setIcon(btn, 'copy');
                        btn.style.backgroundColor = 'var(--interactive-accent)';
                    }, 2000);
                };

                btn.onmouseenter = () => btn.style.transform = 'scale(1.1)';
                btn.onmouseleave = () => btn.style.transform = 'scale(1.0)';
            }
        });
    }

    async copyRichText(view) {
        if (!view) return;

        const markdown = view.data;
        const tempDiv = document.createElement('div');

        await obsidian.MarkdownRenderer.render(this.app, markdown, tempDiv, view.file.path, view);

        this.applyInlineStyles(tempDiv);

        // Use clean plain text (no Markdown symbols) so apps like WeChat get readable text
        const plainText = tempDiv.innerText;

        // Wrap in full HTML document for better cross-app compatibility
        const html = `<!DOCTYPE html><html><head><meta charset="utf-8"><style>body,p,li,ul,ol,h1,h2,h3,h4,h5,h6,td,th,tr,table,span,div,b,i,strong,em{color:#000000;line-height:1.6;mso-line-height-rule:exactly;}body{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Helvetica,Arial,sans-serif;font-size:15px;background-color:white;margin:0;padding:8px;}</style></head><body style="line-height:1.6;mso-line-height-rule:exactly;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Helvetica,Arial,sans-serif;font-size:15px;">${tempDiv.innerHTML}</body></html>`;

        try {
            const blobPlain = new Blob([plainText], { type: "text/plain" });
            const blobHtml = new Blob([html], { type: "text/html" });

            const data = [
                new ClipboardItem({
                    "text/plain": blobPlain,
                    "text/html": blobHtml,
                }),
            ];

            await navigator.clipboard.write(data);
            new obsidian.Notice("✅ Rich text copied! Ready for Outlook/Teams/WeChat.");
        } catch (err) {
            console.error('Rich copy failed:', err);
            new obsidian.Notice("❌ Copy failed. See console for details.");
        }
    }

    applyInlineStyles(container) {
        const fontStack = "-apple-system,BlinkMacSystemFont,'Segoe UI',Helvetica,Arial,sans-serif";
        const monoStack = "ui-monospace,SFMono-Regular,SF Mono,Menlo,Consolas,monospace";
        const lh = "line-height:1.6;mso-line-height-rule:exactly;";
        const lhTight = "line-height:1.3;mso-line-height-rule:exactly;";
        const lhCode = "line-height:1.45;mso-line-height-rule:exactly;";

        const styles = {
            'h1': `font-size:26px;font-weight:700;${lhTight}color:#1a202c;border-bottom:2px solid #edf2f7;padding-bottom:8px;margin:24px 0 12px 0;font-family:${fontStack};`,
            'h2': `font-size:22px;font-weight:600;${lhTight}color:#2d3748;border-bottom:1px solid #edf2f7;padding-bottom:4px;margin:20px 0 10px 0;font-family:${fontStack};`,
            'h3': `font-size:18px;font-weight:600;${lhTight}color:#4a5568;margin:16px 0 8px 0;font-family:${fontStack};`,
            'h4': `font-size:16px;font-weight:600;${lhTight}color:#4a5568;margin:14px 0 6px 0;font-family:${fontStack};`,
            'h5': `font-size:15px;font-weight:600;${lhTight}color:#4a5568;margin:12px 0 4px 0;font-family:${fontStack};`,
            'h6': `font-size:13px;font-weight:600;${lhTight}color:#4a5568;margin:10px 0 4px 0;font-family:${fontStack};`,
            'p': `${lh}color:#333;margin:0 0 8px 0;font-family:${fontStack};font-size:15px;`,
            'ul': `${lh}margin:8px 0;padding-left:24px;font-family:${fontStack};`,
            'ol': `${lh}margin:8px 0;padding-left:24px;font-family:${fontStack};`,
            'li': `${lh}color:#333;margin-bottom:4px;font-family:${fontStack};`,
            'pre': `background-color:#f6f8fa;padding:16px;border-radius:6px;font-family:${monoStack};font-size:14px;${lhCode}overflow:auto;margin:8px 0;border:1px solid #e1e4e8;white-space:pre-wrap;word-wrap:break-word;`,
            'code': `background-color:rgba(175,184,193,0.2);padding:2px 6px;border-radius:4px;font-family:${monoStack};font-size:85%;color:#c7254e;${lhCode}`,
            'table': `border-collapse:collapse;width:100%;margin:8px 0;font-size:14px;font-family:${fontStack};${lh}`,
            'th': `border:1px solid #cbd5e0;padding:8px 13px;background-color:#f7fafc;text-align:left;font-weight:600;${lh}white-space:nowrap;`,
            'td': `border:1px solid #cbd5e0;padding:8px 13px;text-align:left;${lh}`,
            'blockquote': `border-left:4px solid #cbd5e0;padding-left:16px;color:#718096;font-style:italic;margin:8px 0;${lh}`,
            'a': 'color:#3182ce;text-decoration:underline;',
            'strong': 'font-weight:bold;',
            'em': 'font-style:italic;',
            'hr': 'border:none;border-top:1px solid #edf2f7;margin:16px 0;',
            'html': 'background-color:white;',
            'body': `background-color:white;color:black;margin:0;padding:8px;${lh}font-family:${fontStack};font-size:15px;`,
            'span': `${lh}color:#333;font-family:${fontStack};`,
            'div': `${lh}color:#333;font-family:${fontStack};`,
        };

        // Preserve table cell alignment from Obsidian's rendered classes before stripping
        const alignMap = { 'has-text-align-left': 'left', 'has-text-align-center': 'center', 'has-text-align-right': 'right' };
        container.querySelectorAll('th, td').forEach(cell => {
            for (const [cls, align] of Object.entries(alignMap)) {
                if (cell.classList.contains(cls)) {
                    cell.setAttribute('style', (cell.getAttribute('style') || '') + `text-align:${align};`);
                    break;
                }
            }
        });

        // Remove Obsidian-specific UI elements that should not appear in paste output
        container.querySelectorAll(
            '.copy-code-button, .edit-block-button, .tag-pane-tag-count, ' +
            '.markdown-embed-link, .file-embed-title, .collapse-indicator, ' +
            '.callout-fold, [data-callout-fold]'
        ).forEach(el => el.remove());

        // Strip class and id attributes to avoid CSS leakage in target apps
        container.querySelectorAll('[class], [id]').forEach(el => {
            el.removeAttribute('class');
            el.removeAttribute('id');
        });

        // Apply <body> and <html> styles before iterating, so root elements get them
        const htmlEl = container.querySelector('html');
        const bodyEl = container.querySelector('body');
        // The container itself is the root; apply body-level styles to it as fallback
        for (const [tag, style] of Object.entries(styles)) {
            if (tag === 'html' || tag === 'body') continue; // handled separately
            container.querySelectorAll(tag).forEach(el => {
                const existingStyle = el.getAttribute('style') || '';
                el.setAttribute('style', existingStyle + style);
            });
        }
    }

    onunload() {
        document.querySelectorAll('.rich-copy-floating-btn').forEach(btn => btn.remove());
    }
}

module.exports = RichCopyPlugin;