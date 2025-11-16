type LogLevel = 'debug' | 'info' | 'warn' | 'error';

const levelWeight: Record<LogLevel, number> = {
  debug: 10,
  info: 20,
  warn: 30,
  error: 40,
};

const normalizeLevel = (input?: string): LogLevel => {
  const candidate = input?.toLowerCase();
  if (candidate && candidate in levelWeight) {
    return candidate as LogLevel;
  }
  return 'info';
};

export interface LogContext {
  [key: string]: unknown;
}

export class Logger {
  constructor(
    private readonly scope: string,
    private readonly minLevel: LogLevel = 'info'
  ) {}

  child(scope: string): Logger {
    return new Logger(`${this.scope}:${scope}`, this.minLevel);
  }

  debug(message: string, context?: LogContext): void {
    this.emit('debug', message, context);
  }

  info(message: string, context?: LogContext): void {
    this.emit('info', message, context);
  }

  warn(message: string, context?: LogContext): void {
    this.emit('warn', message, context);
  }

  error(message: string, context?: LogContext | Error): void {
    if (context instanceof Error) {
      this.emit('error', message, {
        name: context.name,
        message: context.message,
        stack: context.stack,
      });
      return;
    }
    this.emit('error', message, context);
  }

  private emit(level: LogLevel, message: string, context?: LogContext): void {
    if (levelWeight[level] < levelWeight[this.minLevel]) {
      return;
    }
    const payload = {
      scope: this.scope,
      level,
      message,
      timestamp: new Date().toISOString(),
      ...(context ?? {}),
    };

    const serialized = JSON.stringify(payload);
    switch (level) {
      case 'debug':
        console.debug(serialized);
        break;
      case 'info':
        console.info(serialized);
        break;
      case 'warn':
        console.warn(serialized);
        break;
      case 'error':
        console.error(serialized);
        break;
      default:
        console.log(serialized);
    }
  }
}

const rootLevel = normalizeLevel(process.env.LOG_LEVEL);
export const rootLogger = new Logger('zoom-runner', rootLevel);
export const createLogger = (scope: string): Logger => rootLogger.child(scope);
