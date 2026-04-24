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
        const html = `<!DOCTYPE html><html><head><meta charset="utf-8"></head><body>${tempDiv.innerHTML}</body></html>`;

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
        const styles = {
            'h1': 'font-size: 28px; font-weight: bold; color: #1a202c; border-bottom: 2px solid #edf2f7; padding-bottom: 8px; margin-top: 24px; margin-bottom: 16px; font-family: sans-serif;',
            'h2': 'font-size: 22px; font-weight: bold; color: #2d3748; border-bottom: 1px solid #edf2f7; padding-bottom: 4px; margin-top: 20px; margin-bottom: 12px; font-family: sans-serif;',
            'h3': 'font-size: 18px; font-weight: bold; color: #4a5568; margin-top: 16px; margin-bottom: 8px; font-family: sans-serif;',
            'h4': 'font-size: 16px; font-weight: bold; color: #4a5568; margin-top: 12px; margin-bottom: 8px; font-family: sans-serif;',
            'p': 'margin-bottom: 12px; line-height: 1.6; color: #333; font-family: sans-serif;',
            'ul': 'margin-bottom: 12px; padding-left: 20px; font-family: sans-serif;',
            'ol': 'margin-bottom: 12px; padding-left: 20px; font-family: sans-serif;',
            'li': 'margin-bottom: 4px; line-height: 1.6; color: #333; font-family: sans-serif;',
            'pre': 'background-color: #f6f8fa; padding: 16px; border-radius: 6px; font-family: monospace; font-size: 14px; line-height: 1.45; overflow: auto; margin-bottom: 16px; border: 1px solid #ddd;',
            'code': 'background-color: #f0f0f0; padding: 2px 4px; border-radius: 4px; font-family: monospace; font-size: 14px; color: #c7254e;',
            'table': 'border-collapse: collapse; width: 100%; margin-bottom: 16px; font-family: sans-serif;',
            'th': 'border: 1px solid #cbd5e0; padding: 10px; background-color: #f7fafc; text-align: left; font-weight: bold;',
            'td': 'border: 1px solid #cbd5e0; padding: 10px; text-align: left;',
            'blockquote': 'border-left: 4px solid #cbd5e0; padding-left: 16px; color: #718096; font-style: italic; margin-bottom: 16px;',
            'a': 'color: #3182ce; text-decoration: underline;',
            'strong': 'font-weight: bold;',
            'em': 'font-style: italic;',
            'hr': 'border: none; border-top: 1px solid #edf2f7; margin: 16px 0;'
        };

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

        for (const [tag, style] of Object.entries(styles)) {
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