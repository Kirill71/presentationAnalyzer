/*
 * Copyright (c) 2022 Kyrylo Zharenkov
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in all
 * copies or substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
 * SOFTWARE.
 */

#!/usr/bin/env python

import os
import shutil
import stat
import sys
import argparse

from lxml import etree

os.chmod(sys.argv[0], stat.S_IRWXU | stat.S_IRWXG | stat.S_IRWXO)


class ConsoleColor:
    FOUND = "\033[96m"
    HEADER = '\033[95m'
    OKBLUE = '\033[94m'
    OKGREEN = '\033[92m'
    WARNING = '\033[93m'
    FAIL = '\033[91m'
    ENDC = '\033[0m'
    BOLD = '\033[1m'
    UNDERLINE = '\033[4m'


def prepare_pptx_data(slide_type):
    namespaces = {'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
                  'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}

    def get_slide_type(sld_type):
        return {"slides": "p:sld", "slideLayouts": "p:sldLayout", "slideMasters": "p:sldMaster"}[sld_type]

    base_query = f'/{get_slide_type(slide_type)}/p:cSld/p:spTree/p:pic/p:spPr/'

    blipFill = f'{base_query}a:blipFill'
    solidFill = f'{base_query}/a:solidFill'
    gradientFill = f'{base_query}/a:gradFill'
    patternFill = f'{base_query}/a:pattFill'

    xpathQueries = [blipFill, solidFill, gradientFill, patternFill]

    return xpathQueries, namespaces


def prepare_odp_data(style_name):
    def get_fill_query(style, fill_type, fill_holder):
        baseQuery = f'//style:style[@style:name={style}]'
        return f'{baseQuery}/style:graphic-properties[@draw:fill=\'{fill_type}\']/{fill_holder}'

    blipFill = get_fill_query(style_name, 'bitmap', '@draw:fill-image-name')
    solidFill = get_fill_query(style_name, 'solid', '@draw:fill-color')
    gradientFill = get_fill_query(style_name, 'gradient', '@draw:fill-gradient-name')
    patternFill = get_fill_query(style_name, 'hatch', '@draw:fill-hatch-name')
    xpathQueries = [blipFill, solidFill, gradientFill, patternFill]

    return xpathQueries


def is_xml_contains_xpath_query(xml_tree, xpath_queries, namespaces):
    for query in xpath_queries:
        if len(xml_tree.xpath(query, namespaces=namespaces)) > 0:
            return True

    return False


def unsupported(presentation, message):
    print(f'   {ConsoleColor.WARNING}{presentation}')
    print(f'   {ConsoleColor.BOLD}{ConsoleColor.WARNING}{message}')
    print("   ")


def prepare_args():
    parser = argparse.ArgumentParser()

    parser.add_argument("-i", "--input_dir", type=str, help="Input directory with files for analyzing", required=True)
    parser.add_argument("-o", "--output_dir", type=str, help="Output directory with analyzing results."
                                                             "If it is not specified the result will be created"
                                                             "in the input directory")

    return parser.parse_args()


def get_input_params():
    args = prepare_args()

    if args.input_dir.find(' ') != -1:
        print(ConsoleColor.FAIL + "Invalid input directory name. Input subDir name mustn't contain space "
                                  "symbol. "
                                  "Pls, "
                                  "rename")
        sys.exit(2)

    hasValidOutputPath = args.output_dir is not None and not args.output_dir.isspace()
    return (args.input_dir, args.output_dir) if hasValidOutputPath else (args.input_dir, args.input_dir)


def prepare_path(path):
    return path.replace(' ', '\\ ').replace("(", "\\(").replace(
        ")", "\\)")


def unzip(input_dir_path, presentation):
    tempDirPath = f'{prepare_path(input_dir_path)}/temp'
    full_presentation_path = f'{prepare_path(input_dir_path)}/{prepare_path(presentation)}'
    unzip_command = f'unzip -q -o {full_presentation_path} -d {tempDirPath}'
    os.system(unzip_command)
    return tempDirPath


def is_path_to_slides_exist(path_to_slides, input_dir_path):
    if not os.path.exists(path_to_slides):
        print(path_to_slides)
        print(f'{ConsoleColor.WARNING}The name contains unsupported symbols')
        print(f'{ConsoleColor.WARNING}You should rename containing folder {input_dir_path}')
        print("     ")
        return False

    return True


def save_slides_if_found(presentation, result, input_dir_path, founded_slides):
    if len(founded_slides) > 0:
        fullPresentationPath = f'{input_dir_path}/{presentation}'
        fullPresentationPath.replace("//", "/")
        result[fullPresentationPath] = founded_slides
        print(f'{ConsoleColor.BOLD}{ConsoleColor.FOUND}   FOUND: {presentation}')
        print(ConsoleColor.ENDC)


def analyze_pptx_file(result, input_dir_path, presentation):
    pathToSlides = unzip(input_dir_path, presentation)
    pathToSlides += "/ppt"

    if not is_path_to_slides_exist(pathToSlides, input_dir_path):
        return

    os.chdir(pathToSlides)
    subDirs = ["slides", "slideMasters", "slideLayouts"]
    foundedSlides = set()
    for subDir in subDirs:
        currentPath = f'{pathToSlides}/{subDir}'
        if not os.path.exists(currentPath):
            continue

        os.chdir(currentPath)

        for currentSlide in os.listdir(currentPath):
            if not currentSlide.endswith("_rels"):
                xmlTree = etree.parse(currentSlide)
                xpathQueries, namespaces = prepare_pptx_data(subDir)
                if is_xml_contains_xpath_query(xmlTree, xpathQueries, namespaces):
                    foundedSlides.add(currentSlide[:currentSlide.find('.')])

    save_slides_if_found(presentation, result, input_dir_path, foundedSlides)


def check_fills(slide_number, style_name, xml_tree, namespaces, founded_slides):
    xpathQueries = prepare_odp_data(style_name)

    if is_xml_contains_xpath_query(xml_tree, xpathQueries, namespaces):
        founded_slides.add(f'slide {slide_number}')


def analyze_odp_file(result, input_dir_path, presentation):
    pathToSlides = unzip(input_dir_path, presentation)

    if not is_path_to_slides_exist(pathToSlides, input_dir_path):
        return

    os.chdir(pathToSlides)
    CONTENT = "content.xml"
    contentPath = f'{pathToSlides}/{CONTENT}'
    if not os.path.exists(contentPath):
        print(f'{ConsoleColor.FAIL}Looks like this file is corrupted: {presentation}')
        print(ConsoleColor.ENDC)
        return

    contentXmlTree = etree.parse(contentPath)

    slidesCountXpath = "count(//draw:page)"
    namespaces = {'draw': 'urn:oasis:names:tc:opendocument:xmlns:drawing:1.0',
                  'style': 'urn:oasis:names:tc:opendocument:xmlns:style:1.0',
                  'presentation': 'urn:oasis:names:tc:opendocument:xmlns:presentation:1.0'}

    slidesCount = int(contentXmlTree.xpath(slidesCountXpath, namespaces=namespaces))
    foundedSlides = set()
    for i in range(1, slidesCount + 1):
        framesCountXPath = f'count(//draw:page[{i}]/draw:frame)'
        framesCount = int(contentXmlTree.xpath(framesCountXPath, namespaces=namespaces))
        for j in range(1, framesCount + 1):
            imagesCountXPath = f'count(//draw:page[{i}]/draw:frame[{j}]/draw:image)'
            imageCount = int(contentXmlTree.xpath(imagesCountXPath, namespaces=namespaces))
            if imageCount > 0:
                graphicStyleNameXPath = f'//draw:page[{i}]/draw:frame[{j}]/@draw:style-name'
                graphicStyleName = contentXmlTree.xpath(graphicStyleNameXPath, namespaces=namespaces)
                # The presentation namespace uses if file saved to the odp from MS PoverPoint
                presentationStyleNameXPath = f'//draw:page[{i}]/draw:frame[{j}]/@presentation:style-name'
                presentationStyleName = contentXmlTree.xpath(presentationStyleNameXPath, namespaces=namespaces)
                if len(graphicStyleName) > 0:
                    check_fills(i, graphicStyleName[0], contentXmlTree, namespaces, foundedSlides)
                elif len(presentationStyleName) > 0:
                    check_fills(i, presentationStyleName[0], contentXmlTree, namespaces, foundedSlides)

    save_slides_if_found(presentation, result, input_dir_path, foundedSlides)


def process_file(file_number, presentation, input_dir_path, result):
    file_number += 1
    print(f'{ConsoleColor.OKGREEN}{file_number}. Processing file: {presentation}...')
    if presentation.endswith("pptx") or presentation.endswith("PPTX"):
        analyze_pptx_file(result, input_dir_path, presentation)

    elif presentation.endswith("odp") or presentation.endswith("ODP"):
        analyze_odp_file(result, input_dir_path, presentation)
    elif presentation.endswith("ppt") or presentation.endswith("PPT"):
        unsupported(presentation, "PPT format is binary. The script can't parse it.")
        return file_number
    else:
        unsupported(presentation, "This file type doesn't support.")
        return file_number

    os.chdir(input_dir_path)
    tempDirPath = f'{input_dir_path}/temp'
    if os.path.exists(tempDirPath):
        shutil.rmtree(tempDirPath)

    return file_number


def write_to_file(output_dir_path, result):
    resultFilePath = f'{output_dir_path}/result.txt'
    with open(resultFilePath, 'w+', ) as resultIO:
        if len(result) > 0:
            for key in result.keys():
                resultIO.write(f'Presentation name: {key}\n')
                for slide in sorted(result[key]):
                    resultIO.write(f'      {slide}\n')

                resultIO.write("\n")
        else:
            resultIO.write("This directory doesn't contain files which satisfied the founded conditions.")

    print(f'{ConsoleColor.OKBLUE}Results have been written to: {resultFilePath}')


def main():
    inputDirPath, outputDirPath = get_input_params()

    result = {}

    fileCounter = 0
    for presentation in os.listdir(inputDirPath):
        fullPath = f'{inputDirPath}/{presentation}'
        if os.path.isdir(fullPath):
            for presentationFile in os.listdir(fullPath):
                fileCounter = process_file(fileCounter, presentationFile, fullPath, result)
        elif os.path.isfile(fullPath):
            fileCounter = process_file(fileCounter, presentation, inputDirPath, result)

    write_to_file(outputDirPath, result)

    print(f'{ConsoleColor.OKGREEN}Analyzing was finished successfully!')
    print(ConsoleColor.ENDC)


if __name__ == "__main__":
    main()
