#!/usr/bin/env python
# -*- coding: 'utf-8' -*- 
import loader
def main():
    print("Write querry for all rows\n")	
    loader.querryForAllRows(loader.sheet)
    print("\nRemove duplicates and save to scriptname.sql")	
    loader.writeToScript(loader.sheet)
    loader.removeDuplicates()

if __name__ == '__main__':
    main()
