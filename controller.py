# Controller (polys/mindep only)
# Author :      Nathan Krueger
# Created       5:00 PM 7/16/15
# Last Updated  3:15 PM 6/4/16
# Version       2.61

import UI
import nltk
import xlrd
from openpyxl import *
from nltk.corpus import wordnet

def run()->None:
    """Initializes program"""
    print("""
Please cite as:
Lawrence, J., & Krueger, N. (2016). Bifrost: Bridging linguistic, cognitive and computer science resources (Version Reduced: Polysemy and Depth only). Retrieved from https://github.com/Malthanatos/Bifrost/tree/Reduced--Polysemy-and-Depth-only

This program was developed with support from the University of California Academic Senate Council on Research, Computing, and Libraries (CORCL).

Please contract Joshua Lawrence at jflawren@uci.edu for any of the following papers that have used from data derived from the Bifrost program:

Lawrence, J. F., Hwang, J. K., Hagen, A., & Lin, G. (n.d.). What makes an academic word difficult to know?: Exploring lexical dimensions across novel measures of word knowledge. 

Lawrence, J.F., Hagen, A., Hwang, J. K., Lin, G., & Arne, L. (n.d.). Academic vocabulary and reading comprehension: Exploring the relationships across measures of vocabulary knowledge.

Lawrence, J. F., Lin, G., Jaeggi, S., Krueger, N., Hwang, J. K., & Hagen, A. (n.d.). Polysemy and semantic precision: Standardized semantic measures extracted from Wordnet for 100,000 words in English.

Lawrence, J. F., (n.d.) Semantic precision and polysemy: Key indices of word difficulty and utility for reading.
""")
    print('''Notes: some 2 part words can be analyzed, however, the results
       - of the analysis of such words may be inconsistant depending on
       - whether the input uses a space or an underscore to seperate them''')
    while True:
        interface_data = UI.interface()
        if interface_data[0] == 'quit':
            break
        elif interface_data == (None,None):
            continue
        UI.output_data(collect_data(interface_data))
    return

def collect_data(in_data)->list:
    """collects the data requested"""
    function = in_data[0]
    other = in_data[1]
    if function == 'polys':
        data = polysemy(other)
    if function == 'mindep':
        data = mindepth(other)
    if function == 'pol_min':
        data = polys_mindep(other)
    if function == 'dtree':
        data = depth_tree(other)
    return (data, function)

def pos_redef(pos: str)->str:
    '''converts the 1 letter synset POS into a real word'''
    if pos == 'n':
        return "noun"
    if pos == 'a':
        return "adjective"
    if pos == 's':
        return "satellite adjective"
    if pos == 'r':
        return "adverb"
    if pos == 'v':
        return "verb"

def polysemy(words: [str])->list:
    '''returns a list of polysemy data for a given word set'''
    '''if word_source == 'default':
        file = open('common words.txt')
        words = file.read().splitlines()
    elif word_source == 'manual':
        print("Please enter a string of words seperated only by spaces: ")
        words = input().strip().lower().split()
    #words = [w.lower() for w in word_list]'''
    print("\nGathering data...")
    result = []
    for word in words:
        word_data = [word,0,0,0,0,0,'N/A','N/A','N/A','N/A','N/A']
        word_info = wordnet.synsets(word)
        for synset in word_info:
            #if synset.name().split('.')[0] != word:
                #continue
            if synset.pos() == 'n':
                word_data[1] += 1
            if synset.pos() == 'a':
                word_data[2] += 1
            if synset.pos() == 's':
                word_data[3] += 1
            if synset.pos() == 'r':
                word_data[4] += 1
            if synset.pos() == 'v':
                word_data[5] += 1
        result.append(word_data)
    return result

def mindepth(words: [str])->list:
    '''returns a list of tuples of a word and its min depth'''
    print("\nGathering data...")
    result = []
    for word in words:
        word_data = [word,'N/A','N/A','N/A','N/A','N/A',-1,-1,-1,-1,-1]
        word_info = wordnet.synsets(word)
        for index in range(len(word_info)):
            #if word_info[index].name().split('.')[0] != word:
                #continue
            if word_info[index].pos() == 'n' and word_data[6] == -1:
                word_data[6] = word_info[index].min_depth()
            if word_info[index].pos() == 'a' and word_data[7] == -1:
                word_data[7] = word_info[index].min_depth()
            if word_info[index].pos() == 's' and word_data[8] == -1:
                word_data[8] = word_info[index].min_depth()
            if word_info[index].pos() == 'r' and word_data[9] == -1:
                word_data[9] = word_info[index].min_depth()
            if word_info[index].pos() == 'v' and word_data[10] == -1:
                word_data[10] = word_info[index].min_depth()
        result.append(word_data)
    return result

def polys_mindep(words: [str])->list:
    '''returns a list of lists of words and their depth and polys'''
    #I could shorten this by calling both and merging them, but calling synsets is expensive
    #  and I don't want to do it twice if I can help it
    print("\nGathering data...")
    result = []
    for word in words:
        word_data = [word,0,0,0,0,0,-1,-1,-1,-1,-1]
        word_info = wordnet.synsets(word)
        for index in range(len(word_info)):
            #if word_info[index].name().split('.')[0] != word:
                #continue
            if word_info[index].pos() == 'n':
                if word_data[6] == -1:
                    word_data[6] = word_info[index].min_depth()
                word_data[1] += 1
            if word_info[index].pos() == 'a':
                if word_data[7] == -1:
                    word_data[7] = word_info[index].min_depth()
                word_data[2] += 1
            if word_info[index].pos() == 's':
                if word_data[8] == -1:
                    word_data[8] = word_info[index].min_depth()
                word_data[3] += 1
            if word_info[index].pos() == 'r':
                if word_data[9] == -1:
                    word_data[9] = word_info[index].min_depth()
                word_data[4] += 1
            if word_info[index].pos() == 'v':
                if word_data[10] == -1:
                    word_data[10] = word_info[index].min_depth()
                word_data[5] += 1
        result.append(word_data)
    return result

def depth_tree(word)->str:
    '''returns the word's depth tree'''
    print("Note: only nouns have dtrees, so only noun defintions are displayed")
    print("\nGathering data...")
    #word, defintions, pos, dtrees
    result = ['',[],[],[]]
    result[0] = word
    word_info = wordnet.synsets(word)
    if (len(word_info) > 0):
        word_info = word_info[0]
    else:
        return result
    for synset in wordnet.synsets(word):
        result[1].append(synset.definition())
        result[2].append(synset.pos())
        hyp = lambda w:w.hypernyms()
        result[3].append(synset.tree(hyp))
    return result

if __name__ == '__main__':
    run()
