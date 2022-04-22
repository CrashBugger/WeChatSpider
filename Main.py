from time import sleep

from WeiXin.func import Function

if __name__ == '__main__':
    """此处需要你有一个自己的公众号（注册花5分钟就可以了）
       抓包拿到cookie和get参数token
       如果不急用可以将sleep值调大些"""
    cookie = "你的登录cookie"
    token = "你的token"
    """统计的范围"""
    startTime = "2021-1-1"
    endTime = "2022-4-22"
    func = Function(cookie, token, startTime, endTime)
    pageNum = 0
    while not func.isDone:
        resp = func.getResp(pageNum)
        data = func.getData(resp)
        func.parseAndSaveData(data)
        pageNum += 1
        #sleep最好5秒以上，看任务紧急，太小的话容易被微信暂时冻结半个小时
        sleep(3)
    func.saveData()
