import requests

def get_latest_version():
    """获取最新的版本号"""
    url = "https://orange-nckname.github.io/grades-mover/version.html"
    latest_version = requests.get(url).text
    return float(latest_version)

def is_latest(version_code):
    if version_code == get_latest_version:
        return True
    return False