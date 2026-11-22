/** What the panel's connect view shows: each agent's CLI and SDK as found on this machine. */
import { commandVersion, findOnPath, resolveWorkspacePackage } from './detect.js';

/** @param {import('./provider.js').AgentProvider} provider */
async function providerStatus(cwd, provider) {
  const sdk = resolveWorkspacePackage(cwd, provider.sdkPackage);
  const cliPath = findOnPath(provider.cliCommand);
  const [cliVersion, models] = await Promise.all([cliPath ? commandVersion(cliPath) : null, provider.models()]);
  return {
    id: provider.id,
    name: provider.name,
    sdkPackage: provider.sdkPackage,
    sdkVersion: sdk?.version ?? null,
    cliCommand: provider.cliCommand,
    cliFound: cliPath !== null,
    cliVersion,
    ready: sdk !== null,
    models,
    asksPermission: provider.asksPermission,
  };
}

/** @returns {Promise<import('./types.d.ts').AgentStatus>} */
export async function agentStatus(cwd, providers, packageManager) {
  const list = await Promise.all(Object.values(providers).map((p) => providerStatus(cwd, p)));
  return { enabled: true, packageManager, providers: list };
}
