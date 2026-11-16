import { describe, expect, it } from 'vitest';
import { inline, mdToHtml } from './markdown.ts';

describe('note Markdown', () => {
  it('renders the blocks people write in notes', () => {
    expect(mdToHtml('# 業績\n\n第一行\n第二行\n\n- 甲\n- 乙\n\n1. 一\n2. 二\n\n> 引言\n\n---\n\n```\nx < 1\n```')).toBe(
      '<h1>業績</h1><p>第一行<br>第二行</p><ul><li>甲</li><li>乙</li></ul><ol><li>一</li><li>二</li></ol><blockquote><p>引言</p></blockquote><hr><pre><code>x &lt; 1</code></pre>',
    );
  });

  it('renders inline marks, keeps code literal, and never runs HTML or unsafe links', () => {
    expect(inline('**粗** 與 *斜* 與 `**不是粗體**`')).toBe('<strong>粗</strong> 與 <em>斜</em> 與 <code>**不是粗體**</code>');
    expect(inline('<img src=x onerror=alert(1)>')).toBe('&lt;img src=x onerror=alert(1)&gt;');
    expect(inline('[ok](https://example.com) [bad](javascript:alert(1))')).toBe('<a href="https://example.com" target="_blank" rel="noopener noreferrer">ok</a> bad)');
  });

  it('shows an image only where the note resolves it, and nothing in it is read as Markdown', () => {
    const at = (src: string) => (src.endsWith('.svg') ? `/f?n=${src}` : null);
    expect(inline('![業績 *圖*](a.svg)', at)).toBe('<img class="md-img" src="/f?n=a.svg" alt="業績 *圖*" data-src="a.svg">');
    expect(inline('![x](https://evil.example/x.png)', at)).toBe('x');
    expect(inline('![x](a.svg)')).toBe('x');
    expect(inline('[link](https://a.b) ![p](a.svg)', at)).toContain('<a href="https://a.b"');
    expect(inline('![p" onerror="x](a.svg)', at)).not.toContain('" onerror');
  });
});
