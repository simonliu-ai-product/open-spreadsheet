/**
 * The open-x host protocol: an app that embeds this page (a desktop shell, an
 * editor extension) may inject `window.openxHost` and list what it does. With
 * no host, every feature works as in a plain browser.
 */
type Capability = 'theme' | 'openItem' | 'openHome' | 'guardUnsaved' | 'printToPdf' | 'translucent' | 'agent';

interface OpenxHost {
  protocol: number;
  name: string;
  capabilities: Capability[];
}

const host = (): OpenxHost | null => {
  const h = (window as { openxHost?: OpenxHost }).openxHost;
  return h && h.protocol === 1 && Array.isArray(h.capabilities) ? h : null;
};

export const hostHas = (capability: Capability): boolean => host()?.capabilities.includes(capability) ?? false;
