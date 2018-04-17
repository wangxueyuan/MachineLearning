#!/usr/bin/python
# -*- coding: utf-8 -*-

# __author__ = 'Administrator'
#判断是否回文数
class ParindromeNumber:
    def isPalindrome(self, x):
        """
        :type x: int
        :rtype: bool
        """
        if x<0:
            return False
        num_dic=dict()
        i=0
        while x:
            num_dic.setdefault(i, x%10)
            i+=1
            x //=10
        for j in range((len(num_dic.keys())+1)//2):
            if num_dic.get(j)==num_dic.get(len(num_dic)-j-1):
                continue
            else:
                return False
        return True

if __name__=='__main__':
    parin=ParindromeNumber()

    print(parin.isPalindrome(-2147447412))



