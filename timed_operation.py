from apscheduler.schedulers.blocking import BlockingScheduler
import subprocess
import time

def run_update_script():
    start_time = time.time()
    while True:
        process = subprocess.Popen(['python', './F24看板更新.py'])
        while True:
            if process.poll() is not None:
                break
            elapsed_time = time.time() - start_time
            if elapsed_time > 1200:
                print("F24看板更新.py超过20分钟未完成，重新执行...")
                break
            time.sleep(60)  # 每隔一分钟检查一次是否完成
        if process.poll() is not None:
            break
        start_time = time.time()

    print("F24看板更新.py执行完毕！")

def your_daily_job():
    print("开始运行定时执行程序")
    run_update_script()

scheduler = BlockingScheduler()
scheduler.add_job(your_daily_job, 'cron', hour=0, minute=1)

try:
    scheduler.start()
except (KeyboardInterrupt, SystemExit):
    pass
