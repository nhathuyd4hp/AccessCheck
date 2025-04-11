from app import App
import logging

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    encoding="utf-8",
    datefmt="%Y-%m-%d %H:%M:%S",
    handlers=[
        logging.FileHandler("bot.log", mode="a", encoding="utf-8"),
        logging.StreamHandler(),
    ],
)

if __name__ == "__main__":
    app = App(
        title="Robotic Process Automation",
        geometry="800x600",
        resizable=(False, False),
        icon="robot.ico",
    )
    app.mainloop()
