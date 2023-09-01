class Hello:
    def __init__(self):
        self.name = 'Hello'
    def print(self):
        print(self.name)

def ding():

    h = Hello()
    
    print(h.print())

ding()
