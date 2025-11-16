import { chromium } from 'playwright-extra';
import type { Browser, Frame, Locator, Page } from 'playwright';
import { Command, InvalidArgumentError } from 'commander';
import stealth from 'puppeteer-extra-plugin-stealth';

import { loadConfig, type AppConfig, type ConfigOverrides } from './config.js';
import { createLogger, rootLogger } from './logger.js';

chromium.use(stealth());

type ZoomSurface = Page | Frame;

const signals: NodeJS.Signals[] = ['SIGINT', 'SIGTERM'];

const parseNumber = (value: string): number => {
  const parsed = Number(value);
  if (Number.isNaN(parsed)) {
    throw new InvalidArgumentError('Expected a numeric value');
  }
  return parsed;
};

interface CliOptions {
  url?: string;
  botName?: string;
  message?: string;
  noMonitor?: boolean;
  headless?: boolean;
  headed?: boolean;
  navTimeout?: number;
  nameTimeout?: number;
  lobbyTimeout?: number;
  chatTimeout?: number;
  leaveDelay?: number;
}

const parseCliOverrides = (): ConfigOverrides => {
  const program = new Command();
  program
    .name('zoom-runner')
    .description('Join a Zoom meeting, send a message, and monitor chat for replies')
    .option('--url <url>', 'Zoom meeting URL (overrides ZOOM_URL env)')
    .option('--bot-name <name>', 'Display name used when joining')
    .option('--message <text>', 'Chat message to post once admitted (default: "I\'m in.")')
    .option('--no-monitor', 'Disable message monitoring', false)
    .option('--headless', 'Force headless mode (default)', false)
    .option('--headed', 'Force headed/browser-visible mode', false)
    .option('--nav-timeout <ms>', 'Navigation timeout in ms', parseNumber)
    .option('--name-timeout <ms>', 'Name input wait timeout in ms', parseNumber)
    .option('--lobby-timeout <ms>', 'Lobby wait timeout in ms', parseNumber)
    .option('--chat-timeout <ms>', 'Chat panel wait timeout in ms', parseNumber)
    .option('--leave-delay <ms>', 'Delay before leaving after message in ms', parseNumber);

  const opts = program.parse(process.argv).opts<CliOptions>();

  const headlessOverride = opts.headed ? false : opts.headless ? true : undefined;

  return {
    zoomUrl: opts.url,
    botName: opts.botName,
    messageText: opts.message,
    monitorMessages: opts.noMonitor ? false : undefined,
    headless: headlessOverride,
    navigationTimeoutMs: opts.navTimeout,
    nameInputTimeoutMs: opts.nameTimeout,
    lobbyTimeoutMs: opts.lobbyTimeout,
    chatTimeoutMs: opts.chatTimeout,
    postLeaveDelayMs: opts.leaveDelay,
  };
};

class ZoomRunner {
  private readonly logger = createLogger('zoom-runner');
  private leaveButton?: Locator;
  private currentPage?: Page;

  constructor(private readonly config: AppConfig) {}

  async run(): Promise<void> {
    this.logger.info('Launching Chromium', {
      headless: this.config.headless,
      meetingHost: new URL(this.config.zoomUrl).host,
    });

    const browser = await chromium.launch({
      headless: this.config.headless,
      args: ['--disable-dev-shm-usage', '--no-sandbox', '--disable-setuid-sandbox'],
    });

    const cleanup = this.registerSignalHandlers(browser);

    try {
      const context = await browser.newContext();
      context.setDefaultTimeout(Math.max(this.config.navigationTimeoutMs, this.config.lobbyTimeoutMs));
      const page = await context.newPage();
      this.currentPage = page;
      await this.joinAndChat(page);
    } finally {
      cleanup();
      await browser.close().catch((error) => {
        this.logger.warn('Failed to close browser cleanly', { error });
      });
    }
  }

  private registerSignalHandlers(browser: Browser): () => void {
    const handler = async (signal: NodeJS.Signals) => {
      this.logger.warn('Received termination signal, shutting down', { signal });
      await browser.close().catch(() => undefined);
      process.exit(1);
    };
    signals.forEach((sig) => process.once(sig, handler));
    return () => signals.forEach((sig) => process.removeListener(sig, handler));
  }

  private async joinAndChat(page: Page): Promise<void> {
    try {
      await page.goto(this.config.webClientUrl, {
        waitUntil: 'domcontentloaded',
        timeout: this.config.navigationTimeoutMs,
      });
      await this.acceptCookies(page);
      const surface = await this.resolveSurface(page);
      await this.fillDisplayName(surface);
      await this.submitJoin(surface);
      await this.waitForAdmission(surface);
      await this.sendChatMessage(surface);
      
      if (this.config.monitorMessages) {
        await this.monitorChatMessages(surface);
      }
      
      await this.leaveMeeting(surface);
      this.logger.info('Workflow completed');
    } catch (error) {
      if (page.isClosed()) {
        throw error;
      }
      await page.screenshot({
        path: `artifacts/zoom-runner-failure-${Date.now()}.png`,
        fullPage: true,
      }).catch(() => undefined);
      throw error;
    }
  }

  private async acceptCookies(page: Page): Promise<void> {
    try {
      await page.locator('button:has-text("Accept Cookies")').click({ timeout: 1_000 });
      this.logger.debug('Accepted cookie dialog');
    } catch {
      this.logger.debug('No cookie dialog detected');
    }
  }

  private async resolveSurface(page: Page): Promise<ZoomSurface> {
    const iframeHandle = await page.$('iframe#webclient');
    if (iframeHandle) {
      const frame = await iframeHandle.contentFrame();
      if (frame) {
        return frame;
      }
    }
    const namedFrame = page.frame({ name: 'webclient' });
    return namedFrame ?? page;
  }

  private async fillDisplayName(surface: ZoomSurface): Promise<void> {
    const nameInput = surface.locator('input[type="text"]').first();
    await nameInput.waitFor({ timeout: this.config.nameInputTimeoutMs });
    await nameInput.fill(this.config.botName);
    this.logger.info('Filled display name');
  }

  private async submitJoin(surface: ZoomSurface): Promise<void> {
    await this.dismissBlockingModals(surface, 'before-join');
    
    const candidates = [
      surface.getByRole('button', { name: /^Join(?: Meeting)?$/i }).first(),
      surface.locator('button:has-text("Join")').first(),
      surface.locator('[data-qa="join-meeting"]').first(),
    ];

    // Parallel check for visible buttons
    const visibilityChecks = await Promise.allSettled(
      candidates.map(async (locator, index) => {
        const visible = await locator.isVisible().catch(() => false);
        return { index, locator, visible };
      })
    );

    for (const result of visibilityChecks) {
      if (result.status === 'fulfilled' && result.value.visible) {
        await result.value.locator.click({ timeout: 1_000 });
        this.logger.info('Clicked Join button', { selectorIndex: result.value.index });
        await surface.waitForTimeout(200);
        return;
      }
    }

    // Fallback: try with timeout
    for (const [index, locator] of candidates.entries()) {
      try {
        await locator.click({ timeout: 1_000 });
        this.logger.info('Clicked Join button (fallback)', { selectorIndex: index });
        await surface.waitForTimeout(200);
        return;
      } catch {
        // Continue to next
      }
    }

    this.logger.warn('Join button not found via selectors, sending Enter key fallback');
    await surface.locator('body').press('Enter');
    await surface.waitForTimeout(200);
  }

  private async waitForAdmission(surface: ZoomSurface): Promise<void> {
    this.logger.info('Waiting to be admitted to the meeting', {
      timeoutMs: this.config.lobbyTimeoutMs,
    });

    // Give UI time to settle after clicking Join
    await surface.waitForTimeout(500);

    // Check what state we're in
    const stateIndicators = {
      leaveButton: surface.getByRole('button', { name: /Leave/i }).first(),
      waitingText: surface.locator('text=/waiting|lobby|admitted/i').first(),
      joinAudioButton: surface.locator('[aria-label*="join audio" i]').first(),
      participantsButton: surface.locator('[aria-label*="participant" i]').first(),
      chatButton: surface.locator('[aria-label*="chat" i]').first(),
      videoButton: surface.locator('[aria-label*="video" i]').first(),
    };

    this.logger.debug('Checking meeting state indicators');
    
    const checks = await Promise.allSettled([
      stateIndicators.leaveButton.isVisible().then(v => ({ name: 'leaveButton', visible: v })),
      stateIndicators.waitingText.isVisible().then(v => ({ name: 'waitingText', visible: v })),
      stateIndicators.joinAudioButton.isVisible().then(v => ({ name: 'joinAudioButton', visible: v })),
      stateIndicators.participantsButton.isVisible().then(v => ({ name: 'participantsButton', visible: v })),
      stateIndicators.chatButton.isVisible().then(v => ({ name: 'chatButton', visible: v })),
      stateIndicators.videoButton.isVisible().then(v => ({ name: 'videoButton', visible: v })),
    ]);

    const visibilityState = checks
      .filter((r): r is PromiseFulfilledResult<{ name: string; visible: boolean }> => r.status === 'fulfilled')
      .map(r => r.value);

    this.logger.info('Meeting state after join', { visibilityState });

    // Check if we're already in the meeting (buttons visible)
    const inMeetingButtons = visibilityState.filter(
      s => ['joinAudioButton', 'participantsButton', 'chatButton', 'videoButton'].includes(s.name) && s.visible
    );

    if (inMeetingButtons.length >= 2) {
      this.logger.info('Detected in-meeting UI, skipping lobby wait', {
        visibleButtons: inMeetingButtons.map(b => b.name),
      });
      // Store leave button reference for later
      this.leaveButton = stateIndicators.leaveButton;
      return;
    }

    // Wait for Leave button (indicates admission)
    this.logger.info('Waiting for Leave button to appear (admission indicator)');
    this.leaveButton = stateIndicators.leaveButton;
    
    const pollInterval = 1_000;
    const maxAttempts = Math.ceil(this.config.lobbyTimeoutMs / pollInterval);
    
    for (let attempt = 1; attempt <= maxAttempts; attempt++) {
      try {
        await this.leaveButton.waitFor({ timeout: pollInterval, state: 'visible' });
        this.logger.info('Meeting joined successfully', { attemptNumber: attempt });
        return;
      } catch {
        await this.captureDomSnapshot(surface, `lobby-wait-attempt-${attempt}`);
        this.logger.debug('Still waiting for admission', {
          attempt,
          maxAttempts,
          elapsedMs: attempt * pollInterval,
        });
        
        // Check if we got kicked or meeting ended (quick check)
        const hasError = await surface.locator('text=/ended|closed|removed|invalid|error/i').isVisible().catch(() => false);
        
        if (hasError) {
          await this.captureDebugScreenshot('meeting-error-state');
          throw new Error('Detected meeting error or termination state');
        }
      }
    }

    await this.captureDebugScreenshot('lobby-timeout');
    throw new Error(`Lobby timeout after ${this.config.lobbyTimeoutMs}ms - not admitted to meeting`);
  }

  private async sendChatMessage(surface: ZoomSurface): Promise<void> {
    if (!this.config.messageText) {
      this.logger.info('MESSAGE_TEXT is empty, skipping chat');
      return;
    }

    this.logger.info('Sending chat message');
    await this.dismissBlockingModals(surface, 'before-chat-open');

    const chatButtons = [
      surface.getByRole('button', { name: /^Chat$/i }).first(),
      surface.getByRole('button', { name: /Open Chat/i }).first(),
      surface.getByRole('button', { name: /Chat Panel/i }).first(),
      surface.locator('[aria-label*="chat" i]').first(),
      surface.locator('[data-qa="chat-button"]').first(),
    ];

    // Parallel visibility check for chat buttons
    const chatVisibilityChecks = await Promise.allSettled(
      chatButtons.map(async (locator, index) => {
        const visible = await locator.isVisible().catch(() => false);
        return { index, locator, visible };
      })
    );

    let clickedChatButton = false;
    for (const result of chatVisibilityChecks) {
      if (result.status === 'fulfilled' && result.value.visible) {
        try {
          await result.value.locator.click({ timeout: 500 });
          this.logger.info('Opened chat via button', { selectorIndex: result.value.index });
          clickedChatButton = true;
          break;
        } catch {
          // Try next
        }
      }
    }

    if (!clickedChatButton) {
      this.logger.warn('Chat button not found, trying Alt+H shortcut');
      await surface.locator('body').press('Alt+H');
      await surface.waitForTimeout(300);
    } else {
      // Give chat panel time to fully render
      await surface.waitForTimeout(300);
    }

    const resolvedInput = await this.resolveChatInput(surface);
    if (!resolvedInput) {
      this.logger.warn('Chat input not found, trying activation strategies');
      const activated = await this.tryActivateChatInput(surface);
      if (activated) {
        this.logger.info('Chat input successfully activated');
      } else {
        await this.captureDomSnapshot(surface, 'chat-input-missing');
        await this.captureDebugScreenshot('chat-input-missing');
        throw new Error('Unable to locate or activate chat input after all attempts');
      }
    }

    this.logger.info('Filling chat input');
    const finalInput = resolvedInput || await this.resolveChatInput(surface);
    if (!finalInput) {
      throw new Error('Chat input disappeared after activation');
    }
    await finalInput.fill(this.config.messageText);
    await finalInput.press('Enter');
    this.logger.info('Chat message sent');
  }

  private async monitorChatMessages(surface: ZoomSurface): Promise<void> {
    this.logger.info('Starting continuous message monitoring', {
      botName: this.config.botName,
    });

    const seenMessages = new Set<string>();
    let consecutiveErrors = 0;
    const maxConsecutiveErrors = 10;

    while (true) {
      try {
        // Get all chat messages
        const messages = await surface.evaluate(() => {
          /* eslint-disable @typescript-eslint/no-explicit-any */
          const doc = (globalThis as any).document;
          if (!doc) return [];

          // Try multiple selectors for chat messages
          const selectors = [
            '[data-qa="chat-message"]',
            '[class*="chat-message" i]',
            '[class*="message-item" i]',
            '[id*="message" i]',
            '[role="listitem"]',
          ];

          for (const selector of selectors) {
            const elements = Array.from(doc.querySelectorAll(selector));
            if (elements.length > 0) {
              return elements.map((el: any) => {
                const nameEl = el.querySelector('[class*="sender" i], [class*="name" i], [class*="avatar" i] + *, strong, b');
                const textEl = el.querySelector('[class*="text" i], [class*="content" i], [class*="body" i], p, span');
                
                // Get text, excluding sender name
                let fullText = el.textContent?.trim() || '';
                const senderName = nameEl?.textContent?.trim() || '';
                const messageText = textEl?.textContent?.trim() || fullText;
                
                return {
                  name: senderName,
                  text: messageText,
                  fullText: fullText,
                };
              }).filter((msg: any) => msg.text && msg.text.length > 0);
            }
          }
          return [];
          /* eslint-enable @typescript-eslint/no-explicit-any */
        });

        // Process new messages
        for (const msg of messages) {
          const messageKey = `${msg.name}:${msg.text}`;
          
          if (seenMessages.has(messageKey)) {
            continue;
          }
          
          seenMessages.add(messageKey);
          
          // Skip bot's own messages (check text content for bot replies)
          const isBotOwnMessage = 
            msg.name === this.config.botName || 
            msg.text === 'roger' || 
            msg.text === 'dong' ||
            msg.text === "I'm in.";
          
          if (isBotOwnMessage) {
            this.logger.debug('Skipping bot\'s own message', {
              name: msg.name,
              text: msg.text.substring(0, 50),
            });
            continue;
          }

          this.logger.info('New message received', {
            from: msg.name || 'Unknown',
            text: msg.text.substring(0, 100),
          });

          // Check if message is "quit" command
          const textLower = msg.text.toLowerCase().trim();
          if (textLower === 'quit') {
            this.logger.info('Received quit command, exiting');
            return; // Exit monitoring loop
          }

          // Check if message mentions the bot
          const botNameLower = this.config.botName.toLowerCase();
          const mentionsBot = textLower.includes(botNameLower);
          
          if (mentionsBot) {
            this.logger.info('Bot was mentioned, replying with "dong"', {
              from: msg.name,
              text: msg.text.substring(0, 50),
            });
            await this.sendQuickReply(surface, 'dong');
          } else {
            // Reply "roger" to all other messages
            this.logger.info('Replying with "roger"', {
              from: msg.name,
            });
            await this.sendQuickReply(surface, 'roger');
          }
        }

        consecutiveErrors = 0;
        await surface.waitForTimeout(500); // Check every 500ms for fast responses
      } catch (error) {
        consecutiveErrors++;
        this.logger.debug('Error monitoring messages', {
          error: error instanceof Error ? error.message : String(error),
          consecutiveErrors,
        });

        if (consecutiveErrors >= maxConsecutiveErrors) {
          this.logger.error('Too many consecutive errors, stopping monitoring', {
            consecutiveErrors,
          });
          throw new Error('Message monitoring failed after multiple attempts');
        }

        await surface.waitForTimeout(1_000); // Wait longer on error
      }
    }
  }

  private async sendQuickReply(surface: ZoomSurface, message: string): Promise<void> {
    try {
      // Find chat input quickly
      const input = await this.resolveChatInput(surface);
      if (!input) {
        this.logger.warn('Chat input not found for reply');
        return;
      }

      await input.fill(message);
      await input.press('Enter');
      this.logger.debug('Reply sent successfully', { message });
    } catch (error) {
      this.logger.warn('Failed to send reply', {
        message,
        error: error instanceof Error ? error.message : String(error),
      });
    }
  }

  private async tryActivateChatInput(surface: ZoomSurface): Promise<boolean> {
    this.logger.info('Strategy 1: Clicking inside chat panel area');
    try {
      const chatPanel = surface.locator('[id*="chat" i], [class*="chat" i]').first();
      await chatPanel.click({ timeout: 500, force: true });
      await surface.waitForTimeout(200);
      const input1 = await this.findChatInputCandidates(surface, 'after-panel-click');
      if (input1) {
        this.logger.info('✓ Strategy 1 successful');
        return true;
      }
    } catch (error) {
      this.logger.debug('Strategy 1 failed', { error: error instanceof Error ? error.message : String(error) });
    }

    this.logger.info('Strategy 2: Using Tab key to focus chat input');
    try {
      await surface.locator('body').press('Tab');
      await surface.waitForTimeout(100);
      await surface.locator('body').press('Tab');
      await surface.waitForTimeout(100);
      await surface.locator('body').press('Tab');
      await surface.waitForTimeout(200);
      const input2 = await this.findChatInputCandidates(surface, 'after-tab-keys');
      if (input2) {
        this.logger.info('✓ Strategy 2 successful');
        return true;
      }
    } catch (error) {
      this.logger.debug('Strategy 2 failed', { error: error instanceof Error ? error.message : String(error) });
    }

    this.logger.info('Strategy 3: Clicking chat button again (toggle)');
    try {
      const chatButton = surface.locator('[aria-label*="chat" i]').first();
      await chatButton.click({ timeout: 500 });
      await surface.waitForTimeout(200);
      await chatButton.click({ timeout: 500 });
      await surface.waitForTimeout(300);
      const input3 = await this.findChatInputCandidates(surface, 'after-toggle');
      if (input3) {
        this.logger.info('✓ Strategy 3 successful');
        return true;
      }
    } catch (error) {
      this.logger.debug('Strategy 3 failed', { error: error instanceof Error ? error.message : String(error) });
    }

    this.logger.info('Strategy 4: Waiting longer for lazy loading');
    try {
      await surface.waitForTimeout(1_000);
      const input4 = await this.findChatInputCandidates(surface, 'after-long-wait');
      if (input4) {
        this.logger.info('✓ Strategy 4 successful');
        return true;
      }
    } catch (error) {
      this.logger.debug('Strategy 4 failed', { error: error instanceof Error ? error.message : String(error) });
    }

    this.logger.info('Strategy 5: Searching for ANY visible contenteditable');
    try {
      const anyEditable = surface.locator('[contenteditable="true"]').first();
      const isVisible = await anyEditable.isVisible().catch(() => false);
      if (isVisible) {
        this.logger.info('✓ Strategy 5 found contenteditable', { selector: '[contenteditable="true"]' });
        return true;
      }
    } catch (error) {
      this.logger.debug('Strategy 5 failed', { error: error instanceof Error ? error.message : String(error) });
    }

    this.logger.info('Strategy 6: Checking all visible textareas');
    try {
      const allTextareas = await surface.locator('textarea').all();
      this.logger.debug('Found textareas', { count: allTextareas.length });
      for (const [index, textarea] of allTextareas.entries()) {
        const isVisible = await textarea.isVisible().catch(() => false);
        if (isVisible) {
          const attrs = await textarea.evaluate((el: any) => ({
            id: el.id,
            name: el.name,
            placeholder: el.placeholder,
            className: el.className,
          })).catch(() => null);
          this.logger.debug('Visible textarea found', { index, attrs });
          // Try using first visible textarea that's not obviously for something else
          if (attrs && !attrs.className?.includes('hideme') && !attrs.id?.includes('email')) {
            this.logger.info('✓ Strategy 6 found visible textarea', { index, attrs });
            return true;
          }
        }
      }
    } catch (error) {
      this.logger.debug('Strategy 6 failed', { error: error instanceof Error ? error.message : String(error) });
    }

    this.logger.info('Strategy 7: Trying keyboard shortcut Ctrl+T (chat focus)');
    try {
      await surface.locator('body').press('Control+T');
      await surface.waitForTimeout(300);
      const input7 = await this.findChatInputCandidates(surface, 'after-ctrl-t');
      if (input7) {
        this.logger.info('✓ Strategy 7 successful');
        return true;
      }
    } catch (error) {
      this.logger.debug('Strategy 7 failed', { error: error instanceof Error ? error.message : String(error) });
    }

    this.logger.warn('All activation strategies failed');
    return false;
  }

  private async resolveChatInput(surface: ZoomSurface): Promise<Locator | undefined> {
    this.logger.debug('Searching for chat input in primary surface');
    const primary = await this.findChatInputCandidates(surface, 'surface');
    if (primary) {
      return primary;
    }

    this.logger.debug('Chat input not in primary surface, searching all frames');
    const page = this.pageFromSurface(surface);
    const primaryFrame = this.primaryFrame(surface);
    for (const frame of page.frames()) {
      if (frame === primaryFrame) {
        continue;
      }
      const frameLabel = frame.name() || frame.url();
      const candidate = await this.findChatInputCandidates(frame, `frame:${frameLabel}`);
      if (candidate) {
        return candidate;
      }
    }

    // Capture additional debug info about chat panel state
    this.logger.debug('Analyzing chat panel structure');
    try {
      const chatPanelInfo = await surface.evaluate(() => {
        /* eslint-disable @typescript-eslint/no-explicit-any */
        const doc: any = (globalThis as any).document;
        if (!doc) return { error: 'No document available' };
        
        const chatElements = Array.from(doc.querySelectorAll('[class*="chat" i], [id*="chat" i]')) as any[];
        const editableElements = Array.from(doc.querySelectorAll('[contenteditable="true"]')) as any[];
        const textareas = Array.from(doc.querySelectorAll('textarea')) as any[];
        const inputs = Array.from(doc.querySelectorAll('input[type="text"]')) as any[];
        
        return {
          chatElementsCount: chatElements.length,
          chatElementsSample: chatElements.slice(0, 3).map((el: any) => ({
            tag: el.tagName,
            id: el.id || '',
            classes: el.className || '',
            visible: el.offsetParent !== null,
          })),
          editableCount: editableElements.length,
          textareaCount: textareas.length,
          inputCount: inputs.length,
        };
        /* eslint-enable @typescript-eslint/no-explicit-any */
      });
      this.logger.debug('Chat panel analysis', chatPanelInfo);
    } catch (error) {
      this.logger.debug('Failed to analyze chat panel', {
        message: error instanceof Error ? error.message : String(error),
      });
    }

    return undefined;
  }

  private async findChatInputCandidates(host: ZoomSurface, context: string): Promise<Locator | undefined> {
    const candidates = [
      host.getByRole('textbox', { name: /Chat Input/i }).first(),
      host.getByRole('textbox', { name: /Message/i }).first(),
      host.getByRole('textbox', { name: /Type.*message/i }).first(),
      host.locator('[contenteditable="true"][role="textbox"]').first(),
      host.locator('[contenteditable="true"][data-placeholder]').first(),
      host.locator('[contenteditable="true"].chat-box__chat-textarea').first(),
      host.locator('.chat-box__chat-textarea').first(),
      host.locator('textarea[placeholder*="message" i]').first(),
      host.locator('textarea[placeholder*="type" i]').first(),
      host.locator('input[placeholder*="message" i]').first(),
      host.locator('div[contenteditable="true"]').first(),
    ];

    // Fast parallel visibility check - don't wait for timeout on each selector
    const visibilityChecks = await Promise.allSettled(
      candidates.map(async (locator, index) => {
        const visible = await locator.isVisible().catch(() => false);
        return { index, locator, visible };
      })
    );

    for (const result of visibilityChecks) {
      if (result.status === 'fulfilled' && result.value.visible) {
        this.logger.info('Resolved chat input', { selectorIndex: result.value.index, context });
        return result.value.locator;
      }
    }

    // Fallback: try with short timeout sequentially
    for (const [index, inputLocator] of candidates.entries()) {
      try {
        await inputLocator.waitFor({ timeout: 300, state: 'visible' });
        this.logger.info('Resolved chat input (fallback)', { selectorIndex: index, context });
        return inputLocator;
      } catch {
        // Continue to next selector
      }
    }
    return undefined;
  }

  private pageFromSurface(surface: ZoomSurface): Page {
    if ('context' in surface) {
      return surface;
    }
    return surface.page();
  }

  private primaryFrame(surface: ZoomSurface): Frame {
    if ('page' in surface) {
      return surface;
    }
    return surface.mainFrame();
  }

  private async dismissBlockingModals(surface: ZoomSurface, context: string): Promise<void> {
    const overlayLocator = surface.locator('.ReactModal__Overlay--after-open');
    
    // Quick check first - don't waste time if no overlays
    const initialCount = await overlayLocator.count();
    if (initialCount === 0) {
      return;
    }

    const dismissSelectors = [
      'button:has-text("Got it")',
      'button:has-text("Got It")',
      'button:has-text("OK")',
      'button:has-text("Close")',
      'button:has-text("Continue")',
      '[aria-label="Close"]',
      '[aria-label="close"]',
      '.zm-modal__close',
      '.ReactModalPortal button[aria-label]',
    ];

    this.logger.warn('Blocking overlay detected', { context, overlayCount: initialCount });

    // Try all dismiss buttons in parallel for speed
    await Promise.allSettled(
      dismissSelectors.map(async (selector) => {
        try {
          const dismissButton = surface.locator(selector).first();
          await dismissButton.click({ timeout: 300 });
          this.logger.info('Clicked modal dismiss control', { context, selector });
        } catch {
          // Ignore failures
        }
      })
    );

    // Quick escape attempt
    await surface.locator('body').press('Escape').catch(() => undefined);
    await surface.waitForTimeout(200);

    // Check if overlays remain
    const remaining = await overlayLocator.count();
    if (remaining > 0) {
      this.logger.warn('Blocking overlay still present, force removing', {
        context,
        remaining,
      });
      try {
        const removed = await surface.evaluate(() => {
          const doc = (globalThis as { document?: { querySelectorAll?: (selector: string) => unknown } }).document;
          if (!doc?.querySelectorAll) {
            return 0;
          }
          const elements = doc.querySelectorAll('.ReactModalPortal') as unknown as ArrayLike<{ remove?: () => void }>;
          const portals = Array.from(elements);
          portals.forEach((portal) => portal.remove?.());
          return portals.length;
        });
        this.logger.warn('Force removed ReactModal portals', { context, removed });
      } catch (error) {
        this.logger.error('Failed to force-remove modal portals', error instanceof Error ? error : new Error(String(error)));
      }
      await surface.waitForTimeout(150);
    } else {
      this.logger.info('Blocking overlay dismissed', { context });
    }
  }

  private async captureDomSnapshot(surface: ZoomSurface, reason: string): Promise<void> {
    try {
      const [buttonPreview, textboxPreview] = await Promise.all([
        surface.locator('button').evaluateAll((elements, limit) =>
          elements.slice(0, limit).map((el) => ({
            text: (el.textContent || '').trim().slice(0, 60),
            ariaLabel: el.getAttribute('aria-label') || '',
            dataQa: el.getAttribute('data-qa') || '',
            classes: el.getAttribute('class') || '',
          })),
          12
        ),
        surface
          .locator('[role="textbox"], textarea, input[type="text"], [contenteditable="true"]')
          .evaluateAll((elements, limit) =>
            elements.slice(0, limit).map((el) => ({
              placeholder: el.getAttribute('placeholder') || '',
              ariaLabel: el.getAttribute('aria-label') || '',
              role: el.getAttribute('role') || '',
              text: (el.textContent || '').trim().slice(0, 60),
            })),
            12
          ),
      ]);

      this.logger.debug('DOM snapshot', {
        reason,
        buttons: buttonPreview,
        textboxes: textboxPreview,
      });
    } catch (error) {
      this.logger.debug('Failed to capture DOM snapshot', {
        reason,
        message: error instanceof Error ? error.message : String(error),
      });
    }
  }

  private async captureDebugScreenshot(label: string): Promise<void> {
    if (!this.currentPage) {
      this.logger.debug('Skipping screenshot (no page available)', { label });
      return;
    }
    try {
      const filename = `artifacts/zoom-debug-${label}-${Date.now()}.png`;
      await this.currentPage.screenshot({ path: filename, fullPage: true });
      this.logger.debug('Captured debug screenshot', { label, filename });
    } catch (error) {
      this.logger.debug('Failed to capture screenshot', {
        label,
        message: error instanceof Error ? error.message : String(error),
      });
    }
  }

  private async leaveMeeting(surface: ZoomSurface): Promise<void> {
    this.logger.info('Leaving meeting', { delayMs: this.config.postLeaveDelayMs });
    await surface.waitForTimeout(this.config.postLeaveDelayMs);
    
    // Store reference to leaveButton early
    const leaveButton = this.leaveButton ?? surface.getByRole('button', { name: /Leave/i }).first();
    
    // Try multiple leave strategies quickly
    const leaveStrategies = [
      async () => {
        // Strategy 1: Click stored/found Leave button
        await leaveButton.click({ timeout: 1_000 });
        this.logger.debug('Clicked Leave button');
      },
      async () => {
        // Strategy 2: Try end meeting button (host)
        await surface.locator('button:has-text("End")').first().click({ timeout: 1_000 });
        this.logger.debug('Clicked End button');
      },
      async () => {
        // Strategy 3: Look for any button with leave/end in aria-label
        await surface.locator('[aria-label*="leave" i], [aria-label*="end" i]').first().click({ timeout: 1_000 });
        this.logger.debug('Clicked leave/end via aria-label');
      },
      async () => {
        // Strategy 4: Keyboard shortcut Alt+Q
        await surface.locator('body').press('Alt+Q');
        this.logger.debug('Pressed Alt+Q shortcut');
      },
    ];

    let leftSuccessfully = false;
    for (const [index, strategy] of leaveStrategies.entries()) {
      try {
        await strategy();
        leftSuccessfully = true;
        await surface.waitForTimeout(300);
        break;
      } catch (error) {
        this.logger.debug('Leave strategy failed', {
          strategyIndex: index,
          message: error instanceof Error ? error.message : String(error),
        });
      }
    }

    if (!leftSuccessfully) {
      this.logger.warn('All leave strategies failed, meeting may not be exited cleanly');
      return;
    }

    // Try to confirm leave if dialog appears
    try {
      const confirmButtons = [
        surface.getByRole('button', { name: /Leave Meeting/i }).first(),
        surface.locator('button:has-text("Leave Meeting")').first(),
        surface.locator('button:has-text("Leave")').first(),
      ];

      for (const btn of confirmButtons) {
        try {
          await btn.click({ timeout: 1_000 });
          this.logger.info('Confirmed leave meeting');
          return;
        } catch {
          // Try next
        }
      }
      
      this.logger.debug('No leave confirmation dialog found');
    } catch (error) {
      this.logger.debug('Leave confirmation attempt failed', {
        message: error instanceof Error ? error.message : String(error),
      });
    }
  }
}

async function main() {
  const overrides = parseCliOverrides();
  const config = loadConfig(overrides);
  const runner = new ZoomRunner(config);
  await runner.run();
}

main().catch((error) => {
  const err = error instanceof Error ? error : new Error(String(error));
  rootLogger.error('Fatal error while running Zoom bot', err);
  process.exit(1);
});