import dotenv from 'dotenv';
import { z } from 'zod';

dotenv.config();

const PLACEHOLDER_PATTERN = /(1234567890|example|changeme)/i;

const ZoomUrlSchema = z
  .string()
  .min(1, 'A valid Zoom join URL is required.')
  .url('ZOOM_URL must be a valid https://*.zoom.us link.')
  .refine((value) => !PLACEHOLDER_PATTERN.test(value), {
    message:
      'ZOOM_URL looks like a placeholder. Please provide a real meeting link.',
  });

const ConfigSchema = z.object({
  zoomUrl: ZoomUrlSchema,
  botName: z
    .string()
    .trim()
    .min(1, 'BOT_NAME cannot be empty.')
    .max(120, 'BOT_NAME is too long.')
    .default('Friday BOT'),
  messageText: z
    .string()
    .trim()
    .min(1, 'MESSAGE_TEXT cannot be empty.')
    .max(500, 'MESSAGE_TEXT is too long for the Zoom chat')
    .default("I'm in."),
  monitorMessages: z.coerce.boolean().default(true),
  headless: z.coerce.boolean().default(true),
  navigationTimeoutMs: z.coerce.number().int().positive().default(30_000),
  nameInputTimeoutMs: z.coerce.number().int().positive().default(5_000),
  lobbyTimeoutMs: z.coerce.number().int().positive().default(60_000),
  chatTimeoutMs: z.coerce.number().int().positive().default(1_000),
  postLeaveDelayMs: z.coerce.number().int().positive().default(500),
});

export type ConfigOverrides = Partial<{
  zoomUrl: string;
  botName: string;
  messageText: string;
  monitorMessages: boolean;
  headless: boolean;
  navigationTimeoutMs: number;
  nameInputTimeoutMs: number;
  lobbyTimeoutMs: number;
  chatTimeoutMs: number;
  postLeaveDelayMs: number;
}>;

export type AppConfig = z.infer<typeof ConfigSchema> & {
  webClientUrl: string;
};

export const toWebClientUrl = (url: string): string => {
  const parsed = new URL(url);
  if (parsed.pathname.includes('/wc/join/')) {
    return parsed.toString();
  }

  if (parsed.pathname.includes('/j/')) {
    parsed.pathname = parsed.pathname.replace('/j/', '/wc/join/');
  }
  return parsed.toString();
};

export function loadConfig(overrides: ConfigOverrides = {}): AppConfig {
  const raw: Record<string, unknown> = {
    zoomUrl: overrides.zoomUrl ?? process.env.ZOOM_URL,
    botName: overrides.botName ?? process.env.BOT_NAME,
    messageText: overrides.messageText ?? process.env.MESSAGE_TEXT,
    headless: overrides.headless ?? process.env.HEADLESS,
    navigationTimeoutMs:
      overrides.navigationTimeoutMs ?? process.env.NAVIGATION_TIMEOUT_MS,
    nameInputTimeoutMs:
      overrides.nameInputTimeoutMs ?? process.env.NAME_INPUT_TIMEOUT_MS,
    lobbyTimeoutMs: overrides.lobbyTimeoutMs ?? process.env.LOBBY_TIMEOUT_MS,
    chatTimeoutMs: overrides.chatTimeoutMs ?? process.env.CHAT_TIMEOUT_MS,
    postLeaveDelayMs:
      overrides.postLeaveDelayMs ?? process.env.POST_LEAVE_DELAY_MS,
  };

  const parsed = ConfigSchema.parse(raw);

  return {
    ...parsed,
    webClientUrl: toWebClientUrl(parsed.zoomUrl),
  };
}
