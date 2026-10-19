/* pm2 process definition — diary reader HTTP server.

   Keeps the local diary page (DiaryReader.html + /api/diary) running
   continuously. serve.py loads local/.env itself (port, DB path), so no
   env is needed here.

     pm2 start ecosystem.config.cjs     # start + daemonize
     pm2 save                            # persist so it survives reboot
     pm2 logs diary-server               # tail output

   The one-shot sync scripts (pull_diary.py / enrich_diary.py) are meant to
   run on a schedule and exit, so they are NOT managed here — schedule them
   with launchd/cron instead (pm2 cron is awkward for run-and-exit jobs).
*/
module.exports = {
  apps: [
    {
      name: 'diary-server',
      script: './serve.py',
      cwd: __dirname,
      interpreter: 'python3',
      autorestart: true,
      max_restarts: 20,
      restart_delay: 2000,
      exp_backoff_restart_delay: 100,
    },
  ],
};
