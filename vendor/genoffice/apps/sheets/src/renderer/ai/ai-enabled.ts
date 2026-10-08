/// Embedders that bring their own AI (or none) build with
/// VITE_GENOFFICE_AI=false: the AI panel, its ribbon group and the selection
/// prompt are then left out.
export const AI_ENABLED = import.meta.env.VITE_GENOFFICE_AI !== 'false'
