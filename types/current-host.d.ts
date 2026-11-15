// Additions in the Zana 2.6.1 source SDK, pending the CLI ambient facade refresh.
declare module '@zana-ai/zcc-plugin-sdk/app' {
 export interface PluginThreadPanelProps {pluginId:string;threadId:string;projectId?:string;params:unknown}
 export interface PluginFileOpenerSource {kind:'workspace'|'host'|'thread-storage';threadId:string|null;environmentId:string|null;projectId:string|null;hostId?:string}
 export interface PluginFileOpenerProps {pluginId:string;path:string;source:PluginFileOpenerSource}
 export interface PluginComposerApi {addQuote(text:string):void;focus():void}
}
