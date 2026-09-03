import { spawn } from 'node:child_process';
import { fileURLToPath } from 'node:url';

export const inject = [];

export function apply(ctx) {
  ctx.effect(() => {
    const serverPath = fileURLToPath(new URL('../server/export-file.mjs', import.meta.url));
    const child = spawn(process.execPath, [serverPath], {
      stdio: 'inherit',
      env: process.env,
    });

    child.once('error', (error) => {
      ctx.logger?.error?.('Failed to start the Univer file export service: %s', error.message);
    });

    return () => {
      if (!child.killed) child.kill('SIGTERM');
    };
  }, 'dsh-univer-file-export.server');
}
