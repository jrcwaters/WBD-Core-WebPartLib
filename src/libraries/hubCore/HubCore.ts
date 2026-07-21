/**
 * SPFx Library component anchor for @wbd/hub-core.
 *
 * This class exists so the solution has a registered client-side Library
 * component (see HubCore.manifest.json). The shared data layer the web parts
 * actually consume is re-exported from ../../index.ts.
 */
export class HubCore {
  public name(): string {
    return 'HubCore';
  }
}
