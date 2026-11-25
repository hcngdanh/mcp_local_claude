

// Tools are read-only unless they declare otherwise. Clients (e.g. Claude
// Desktop) use these hints to decide when to ask the user for approval.
export const READ_ONLY_ANNOTATIONS = {
  readOnlyHint: true,
  destructiveHint: false,
  idempotentHint: true,
  openWorldHint: true,
};

export function registerAliases(server, name, config, handler, aliases = []) {
  config = { ...config, annotations: { ...READ_ONLY_ANNOTATIONS, ...(config.annotations || {}) } };
  server.registerTool(name, config, handler);
  for (const alias of aliases) {
    server.registerTool(alias, config, handler);
  }
}


export function errorResult(err) {
  return {
    isError: true,
    content: [{ type: "text", text: String(err?.message || err) }],
  };
}
