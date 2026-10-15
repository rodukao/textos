import type { Config } from 'superdoc';

/**
 * Configuração do SuperDoc compartilhada entre o app e a bancada de fidelidade
 * (bench/harness), para que a bancada meça exatamente o que o app faz.
 */
export const baseSuperDocConfig = {
  // O SuperDoc envia um evento a cada documento aberto, por padrão.
  // Os documentos são do usuário: nada sai da máquina.
  telemetry: { enabled: false },
} satisfies Partial<Config>;

export const DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
