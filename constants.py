#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os

argparse = {
    "description": "Script python qui génère un fichier Excel pour trouver le meilleur investissement BNP Paribas.",
    "items": [
        {
            "name": "Common Options",
            "items": [
                {
                    "name": "debug",
                    "short": "d",
                    "description": "Enable debugging",
                    "default": False
                }
            ]
        },
        {
            "name": "Input",
            "items": [
                {
                    "name": "country",
                    "short": "c",
                    "description": "Country (default is %(default)s)",
                    "enum":    [
                        "AUT", "HRV", "FIN", "HUN", "LIE", "PRT", "SWE", "BHR", "CYP", "FRA", "IRL", "LUX", "SVK", "CHE", "BEL", "CZE", "DEU", "ITA", "NOR", "SVN", "NLD", "CHL", "DNK", "GRC", "JER", "POL", "SPA", "GBR", "AUS", "MAC", "TWN", "HKG", "MYR", "IDN", "SGP", "JPN", "KOR", "BRA", "PER", "USA"
                    ],
                    "default": "FRA"
                },
                {
                    "name": "language",
                    "short": "l",
                    "description": "Language (default is %(default)s)",
                    "enum":    [
                        "FRE", "ENG"
                    ],
                    "default": "FRE"
                },
                {
                    "name": "isin",
                    "short": "i",
                    "description": "List of comma separated funds or path to a fund list file (all BNP Paribas funds by default)",
                    "default": None
                },
                {
                    "name": "favorites",
                    "short": "f",
                    "description": "Path to a comma separated list file (csv) with some ISIN to mark as favorite (default is %(default)s)",
                    "default": "favorites.csv"
                },
                {
                    "name": "type",
                    "short": "t",
                    "description": "Type of investor (default is %(default)s)",
                    "enum":    [
                        "Private investor",
                        "Institutional investor or Financial intermediaries"
                    ],
                    "default": "Private investor"
                }
            ]
        },
        {
            "name": "Output",
            "items": [
                {
                    "name": "file",
                    "short": "o",
                    "description": "Output Excel file (default is %(default)s)",
                    "default": f"{os.getcwd()}/arbitrage.xlsx"
                }
            ]
        }
    ]
}

type_to_api_prefix = {
    "Private investor": "IP_FR-IND",
    "Institutional investor or Financial intermediaries": "PV_FR-FSE",
}

type_to_website_prefix = {
    "Private investor": "individuel",
    "Institutional investor or Financial intermediaries": "intermediaires",
}

favorites_main_key = "isin"

more_details_domain = "https://www.quantalys.com"
more_details_cookie = "UQY4IPIWOASM4GQGUWJBCHPU4VEQPCGBKDRKFXCHSXJIWTIOYPQKCY2NFOO4RZ7LAU6NNSQQX5UVQJT767P677SOKY3SEW74PBUDHBEQEWH4E===;"
more_details_user_agent = "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/138.0.0.0 Safari/537.36"

website_domain = "https://www.bnpparibas-am.com"

api_endpoint = "https://api.bnpparibas-am.com"

currency_code_to_symbol = {
    "EUR": "€",
    "USD": "$",
    "GBP": "£",
    "JPY": "¥",
    "CAD": "$",
    "AUD": "$",
    "CHF": "Fr",
    "CNY": "¥",
    "SEK": "kr",
    "NZD": "$"
}

known_fee_keys = [
    "at_launch_ongoing_charges",
    "estimated_ongoing_charges",
    "maximum_conversion_rate",
    "maximum_management_fees",
    "maximum_redemption_fixed_fees_acquired",
    "maximum_redemption_fixed_fees",
    "maximum_subscription_fixed_fees_acquired",
    "maximum_subscription_fixed_fees",
    "perf_benchmark_spread",
    "real_ongoing_charges",
    "redemption_fixed_fees_acquired",
    "total_redemption_fees",
    "total_subscription_fees",
]

breakdowns_mapping = {
    "countries": [
        "FUNDSHEET_HOLDINGS_TITLE_BY_COUNTRY",
        "FUNDSHEET_HOLDINGS_TITLE_BY_COUNTRY_BENCH"
    ],
    "currencies": [
        "FUNDSHEET_HOLDINGS_TITLE_BY_CURRENCY"
    ],
    "holdings": [
        "FUNDSHEET_HOLDINGS_MAIN_HOLDINGS"
    ],
    "sectors": [
        "FUNDSHEET_HOLDINGS_TITLE_BY_SECTOR_BENCH",
        "FUNDSHEET_HOLDINGS_TITLE_BY_SECTOR",
        "FUNDSHEET_HOLDINGS_MAQS_TYPE",
    ]
}

breakdowns_exclude = [
    "FUNDSHEET_HOLDINGS_TITLE_BY_RATINGS"
]

worksheet = {
    "color":  "1072BA",
    "title":  "Assets"
}

column_mapping = [
    {
        "name": "",
        "items": [
            {
                "ref": "favorite",
                "name": "⭐",
                "width": 3
            },
            {
                "ref": "isin",
                "name": "ISIN",
                "width": 14
            }
        ]
    },
    {
        "name": "Présentation",
        "items": [
            {
                "ref": "asset_class",
                "name": "Classe\nd'actif"
            },
            {
                "ref": "asset_region_class",
                "name": "Région de\ndiversification",
                "conditional-formatting": {
                    "fill-mapping": {
                        "Amérique du Nord": "A13320",
                        "Asie-Pacifique": "560F75",
                        "Europe": "336BD4",
                        "Eurozone": "0E307D"
                    }
                }
            },
            {
                "ref": "fundshare_id",
                "name": "Id\nFond",
                "width": 6
            },
            {
                "ref": "legal_name",
                "name": "Nom Légal"
            },
            {
                "ref": "legal_form",
                "name": "Forme\nJuridique"
            },
            {
                "ref": "creation_date",
                "name": "Date de\ncréation"
            },
            {
                "ref": "share_type",
                "name": "Type de parts"
            },
            {
                "ref": "share_size",
                "name": "Actif total\nde la part"
            },
            {
                "ref": "share_vl",
                "name": "Valeur\nliquidative",
                "width": 6
            },
            {
                "ref": "currency",
                "name": "Devise",
                "width": 6,
                "conditional-formatting": {
                    "fill-mapping": {
                        "Dollar": "A13320"
                    }
                }
            },
            {
                "ref": "base_index",
                "name": "Indice de référence"
            },
            {
                "ref": "sri_risk",
                "name": "Indicateur\nRisque",
                "width": 6,
                "conditional-formatting": {
                    "fill-percentile": {
                        "start_color": "005E23",
                        "mid_color": "946A00",
                        "end_color": "610000"
                    }
                }
            },
            {
                "ref": "morning_star",
                "name": "Morning\nStar",
                "width": 6,
                "conditional-formatting": {
                    "fill-percentile": {
                        "start_color": "610000",
                        "mid_color": "946A00",
                        "end_color": "005E23"
                    }
                }
            },
            {
                "ref": "q_notation",
                "name": "Notation\nQuantalys",
                "width": 6,
                "conditional-formatting": {
                    "fill-percentile": {
                        "start_color": "610000",
                        "mid_color": "946A00",
                        "end_color": "005E23"
                    }
                }
            },
            {
                "ref": "pea",
                "name": "Eligible\nPEA",
                "width": 6
            },
            {
                "ref": "policy",
                "name": "Politique d'investissement",
                "width": 40
            },
            {
                "ref": "source_details",
                "name": "Source",
                "width": 6
            }
        ]
    },
    {
        "name": "Performances",
        "items": [
            {
                "ref": "perf_cumulated",
                "name": "Perf\ncumulée\n5 ans",
                "width": 6,
                "conditional-formatting": {
                    "fill-percentile": {
                        "start_color": "610000",
                        "mid_color": "946A00",
                        "end_color": "005E23"
                    }
                }
            },
            {
                "ref": "perf_cumulated_diff",
                "name": "Diff\nBase",
                "width": 6,
                "conditional-formatting": {
                    "fill-percentile": {
                        "start_color": "610000",
                        "mid_color": "946A00",
                        "end_color": "005E23"
                    }
                }
            },
            {
                "ref": "volatility",
                "name": "Volatilité",
                "width": 6,
                "conditional-formatting": {
                    "fill-percentile": {
                        "start_color": "005E23",
                        "mid_color": "946A00",
                        "end_color": "610000"
                    }
                }
            },
            {
                "ref": "sharpe_ratio",
                "name": "Ratio de\nSharpe",
                "width": 6,
                "conditional-formatting": {
                    "fill-percentile": {
                        "start_color": "610000",
                        "mid_color": "946A00",
                        "end_color": "005E23"
                    }
                }
            },
            {
                "ref": "dic_details",
                "name": "Document\nd'Informations\nClés",
                "width": 6
            },
            {
                "ref": "more_details",
                "name": "Détails",
                "width": 6
            }
        ]
    },
    {
        "name": "Rendement Scénarios 5 ans",
        "items": [
            {
                "ref": "scenario_stressed",
                "name": "Tensions",
                "width": 6,
                "conditional-formatting": {
                    "fill-percentile": {
                        "start_color": "610000",
                        "mid_color": "946A00",
                        "end_color": "005E23"
                    }
                }
            },
            {
                "ref": "scenario_unfavorable",
                "name": "Défavorable",
                "width": 6,
                "conditional-formatting": {
                    "fill-percentile": {
                        "start_color": "610000",
                        "mid_color": "946A00",
                        "end_color": "005E23"
                    }
                }
            },
            {
                "ref": "scenario_moderate",
                "name": "Intermédiaire",
                "width": 6,
                "conditional-formatting": {
                    "fill-percentile": {
                        "start_color": "610000",
                        "mid_color": "946A00",
                        "end_color": "005E23"
                    }
                }
            },
            {
                "ref": "scenario_favorable",
                "name": "Favorable",
                "width": 6,
                "conditional-formatting": {
                    "fill-percentile": {
                        "start_color": "610000",
                        "mid_color": "946A00",
                        "end_color": "005E23"
                    }
                }
            }
        ]
    },
    {
        "name": "Portefeuille",
        "items": [
            {
                "ref": "portfolio_holdings",
                "name": "Principales\nHoldings"
            },
            {
                "ref": "portfolio_currencies",
                "name": "Devises"
            },
            {
                "ref": "portfolio_sectors",
                "name": "Secteurs"
            },
            {
                "ref": "portfolio_countries",
                "name": "Pays"
            }
        ]
    },
    {
        "name": "Frais",
        "items": [
            {
                "ref": "fee_conversion_rate",
                "name": "Coûts de\nconversion",
                "width": 6,
                "conditional-formatting": {
                    "fill-percentile": {
                        "start_color": "005E23",
                        "mid_color": "946A00",
                        "end_color": "610000"
                    }
                }
            },
            {
                "ref": "fee_ongoing_charges",
                "name": "Frais courants\nestimés",
                "width": 6,
                "conditional-formatting": {
                    "fill-percentile": {
                        "start_color": "005E23",
                        "mid_color": "946A00",
                        "end_color": "610000"
                    }
                }
            },
            {
                "ref": "fee_maximum_subscription",
                "name": "Frais\nd'entrée max",
                "width": 6,
                "conditional-formatting": {
                    "fill-percentile": {
                        "start_color": "005E23",
                        "mid_color": "946A00",
                        "end_color": "610000"
                    }
                }
            },
            {
                "ref": "fee_maximum_redemption",
                "name": "Frais de\nsortie max",
                "width": 6,
                "conditional-formatting": {
                    "fill-percentile": {
                        "start_color": "005E23",
                        "mid_color": "946A00",
                        "end_color": "610000"
                    }
                }
            },
            {
                "ref": "fee_real_ongoing",
                "name": "Frais courants\nréels",
                "width": 6,
                "conditional-formatting": {
                    "fill-percentile": {
                        "start_color": "005E23",
                        "mid_color": "946A00",
                        "end_color": "610000"
                    }
                }
            },
            {
                "ref": "fee_redemption_acquired",
                "name": "Commissions de rachat\nacquises au fonds",
                "width": 6,
                "conditional-formatting": {
                    "fill-percentile": {
                        "start_color": "005E23",
                        "mid_color": "946A00",
                        "end_color": "610000"
                    }
                }
            },
            {
                "ref": "fee_maximum_management",
                "name": "Commission de\ngestion max",
                "width": 6,
                "conditional-formatting": {
                    "fill-percentile": {
                        "start_color": "005E23",
                        "mid_color": "946A00",
                        "end_color": "610000"
                    }
                }
            }
        ]
    }
]
