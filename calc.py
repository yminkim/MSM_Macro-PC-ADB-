class FourCal:
    def setData(self, first, second):
        self.first = first
        self.second = second

    def add(self):
        result = self.first + self.second
        return result
    
    def mul(self):
        result = self.first * self.second
        return result
    def div(self):
        result = self.first / self.second
        return result
    
    def sub(self):
        result = self.first - self.second
        return result
        
a = FourCal()
a.setData(8,4)
print(a.add())
print(a.mul())
print(a.div())
print(a.sub())
 
#객체를 사용한 계산기...'-` 
