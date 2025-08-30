#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import argparse
import csv
import datetime
import json
import os
import requests
from stdnum import isin
from pylogger_unified import logger as pylogger_unified
import constants

logger = pylogger_unified.init_logger(json_formatter=False, enable_gi=False)


def parse_args():
    parser = argparse.ArgumentParser(description=constants.argparse["description"])
    for group in constants.argparse["items"]:
        arg_group = parser.add_argument_group(group["name"])
        for item in group["items"]:
            arg_dict = {
                "help": item["description"]
            }
            if "enum" in item:
                arg_dict["choices"] = item["enum"]
            if isinstance(item["default"], str):
                arg_dict["default"] = item["default"]
            elif isinstance(item["default"], bool):
                arg_dict["action"] = "store_" + str(not item["default"]).lower()
            arg_group.add_argument(f"-{item['short']}", f"--{item['name']}", **arg_dict)

    return parser.parse_args()


def merge_lists_deduped(list1, list2):
    if not isinstance(list1, list) or not isinstance(list2, list):
        raise TypeError
    merged = list1[:]
    seen = set(list1)
    for x in list2:
        if x not in seen:
            merged.append(x)
            seen.add(x)
    return merged


def read_file_txt(file_path):
    # only funds from list selected
    try:
        with open(file_path, "r", encoding="utf-8") as file:
            funds = file.readlines()  # Read all lines into a list
        # Remove trailing newlines from each line.  Important for consistency.
        funds = [line.rstrip("\n\r") for line in funds if not line.startswith("#")]
    except FileNotFoundError:
        logger.error(f"Error: File not found at {file_path}")
        raise
    except Exception as e:
        logger.error(f"An error occurred: {e}")
        raise
    return funds


def read_file_csv(file_path):
    data = {}
    try:
        with open(file_path, "r", newline="", encoding="utf-8") as file:
            csv_reader = csv.DictReader(file)
            for row in csv_reader:
                lowercase_row = {k.lower(): v for k, v in row.items()}
                main_key = lowercase_row.pop(constants.favorites_main_key)
                data[main_key] = lowercase_row
    except FileNotFoundError:
        logger.error(f"Error: File not found at {file_path}")
        raise
    except Exception as e:
        logger.error(f"An error occurred: {e}")
        raise
    return data


def check_args(args):
    if not os.path.exists(os.path.dirname(args.file)):
        os.makedirs(os.path.dirname(args.file))  # Create the directory and any necessary parent directories
        logger.warning(f"Directory {os.path.dirname(args.file)} created.")

    if args.isin is not None:
        if all([isin.is_valid(item) for item in args.isin.split(",")]):
            # list provided from command line
            args.isin = args.isin.split(",")
            logger.warning("A comma separated list of funds has been provided in the script arguments")
        elif os.access(args.isin, os.R_OK) or os.access(f"{os.getcwd()}/{args.isin}", os.R_OK):
            isin_path = args.isin
            if os.access(f"{os.getcwd()}/{args.isin}", os.R_OK):
                isin_path = f"{os.getcwd()}/{args.isin}"
            res = read_file_txt(isin_path)
            if not all([isin.is_valid(item) for item in res]):
                raise OSError(f"File {isin_path} is not a valid list of ISIN")
            logger.warning(f"File {isin_path} containing list of funds has been provided")
            args.isin = res
        else:
            raise OSError(f"File {args.isin} is not a readable file or not a valid list of ISIN")

    if os.access(args.favorites, os.R_OK) or os.access(f"{os.getcwd()}/{args.favorites}", os.R_OK):
        logger.warning("A csv file with favorite funds has been found")
        args.favorites = read_file_csv(args.favorites)
        args.isin = merge_lists_deduped(args.isin, list(args.favorites.keys()))
    else:
        args.favorites = {}

    if not os.access(os.path.dirname(args.file), os.W_OK):
        raise OSError(f"Directory {os.path.dirname(args.file)} is not writable !")

    if os.path.splitext(os.path.basename(args.file))[1] != ".xlsx":
        raise OSError(f"File {os.path.basename(args.file)} must have xlsx extension !")


def request_data(url, method="GET", data=None, headers=None, cookies=None):
    try:
        request_method = getattr(requests, method.lower())
        if method.upper() == "GET":
            response = request_method(url, headers=headers, cookies=cookies)
        else:
            response = request_method(url, headers=headers, cookies=cookies, data=data)

        # Raise an exception for bad status codes (4xx or 5xx)
        response.raise_for_status()

        # Attempt to parse the JSON response
        data = response.json()
        return data

    except requests.exceptions.RequestException as e:
        # More informative error messages
        logger.error(f"Error making API request: {e}")
        raise

    except json.JSONDecodeError as e:
        logger.error(f"Error decoding JSON: {e}")
        # Show the actual response content that failed to parse
        logger.error(f"Response content: {response.content}")
        raise


def remove_invalid_xml_chars(text):
    if text is None:
        return None
    if not isinstance(text, str):
        return text
    return "".join(c for c in text if ord(c) >= 32 or c in ("\t", "\n", "\r"))


def join_h(lst):
    if len(lst) == 0:
        return ""
    if len(lst) == 1:
        return lst[0]
    last_item = lst.pop()
    return " and ".join([", ".join(lst), last_item])


def get_utc_time():
    utc_now = datetime.datetime.utcnow()
    return utc_now.strftime("%Y-%m-%d %H:%M:%S UTC")
