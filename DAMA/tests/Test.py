from dataclasses import dataclass
from functools import wraps
from check_test import check_method_input, check_method_input_test
from ApplyToDecorator_test import for_all_methods
from typing import TypeVar


Self = TypeVar('Self', bound='yolo')
@dataclass
class bob:
    age:int 
    family_name:str = ''

@for_all_methods(check_method_input_test(("opt",)), "")
class yolo (object):
    def __init__(self, a:int) -> None:
        self.a=a
    
    def func_1 (self, b:int) -> int:
        #print('f1')
        return self.a + b
    
    def func_2 (self) -> int:
        #print('f2')
        return 2*self.a
    
    def func_3 (self,*,opt) -> None:
        print("Le caca des oiseaux c'est caca")
    
    def func_4 (self,*, age:int=8, bobi:str|int='bob') -> bob:
        #print('f4')
        bob_1 = bob(age, bobi)
        return bob_1
    
    def func_5 (self, b:bob):
        b.age += 1
        print('func_5: ', b)

class yolo_2(object):
    def __init__() -> None:
        pass

    def proute(a:int, b:int):
        return range(1, 2).stop + 10

def main():
    print('func_1: ', yolo(2).func_1(3))
    print('func_2: ', yolo(2).func_2())
    yolo(2).func_3()
    bobi_1 = yolo(2).func_4(bobi = 'booba')
    yolo(2).func_5(bobi_1)
    print('a+b: ', yolo_2.proute(2,3))

if __name__ == "__main__":
    main()

