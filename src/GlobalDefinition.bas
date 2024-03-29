Attribute VB_Name = "GlobalDefinition"
' 定数
    ' コース数
    Public Const TRACK_NUM = 96
    ' レース数
    Public Const RACE_NUM = 12
    ' シート保護パスワード
    Public Const PROTECT_PASSWORD = "MK8DX"
    ' データカラム数
    Public Const DATA_COLS = 11
    ' 登録キー桁数
    Public Const REGIST_KEY_MAX = 999999
    ' データテーブル名
    Public Const DATA_TABLE_NAME = "テーブル1"
    ' ピボットテーブル名
    Public Const GRAPH_PIVOT_TABLE_NAME = "ピボットテーブル1"
    ' 最大知識数
    Public Const MAX_KNOWLEDGE = 4

' シート名
    Public Const REGIST_DATA = "RegistData"
    Public Const STORAGE = "Storage"
    Public Const SETTINGS = "Settings"
    Public Const DATA = "Data"
    Public Const GRAPH = "Graph"
    Public Const TRACK_MASTER = "TrackMaster"
    Public Const KNOWLEDGE_MASTER = "KnowledgeMaster"
    Public Const CUP_MASTER = "CupMaster"
    Public Const VERSION_MASTER = "VersionMaster"
    Public Const LOUNGE_TIER_MASTER = "LoungeTierMaster"
    Public Const FORMAT_MASTER = "FormatMaster"
    Public Const POINT_MASTER = "PointMaster"
    Public Const LANGUAGE_MASTER = "LanguageMaster"
    
' 文言
    Public Const SELECT_TRACK_JP = "コースを選択"
    Public Const SELECT_TRACK_EN = "Select Track"
    Public Const UNSELECT_JP = "未選択"
    Public Const UNSELECT_EN = "Unselected"
    
' ピボットテーブル関連
    Public Const PIVOT_FILTER_NAME_1 = "模擬tier / Match tier"
    Public Const PIVOT_FILTER_NAME_2 = "形式 / Format"
    Public Const PIVOT_FILTER_NAME_3 = "スタート順位 / Starting rank"
    Public Const PIVOT_FILTER_NAME_4 = "備考 / Detail"
    Public Const PIVOT_ROW_NAME = "コース名 / Track name"
    Public Const PIVOT_COL_NAME_1 = "Ave. points"
    Public Const PIVOT_COL_NAME_2 = "Ave. rank"
    Public Const PIVOT_COL_NAME_3 = "回数 / Times"

' コース画像関連
    ' ディレクトリ
    Public Const TRACK_IMG_DIR = "\track_img\"
    ' 配置位置x
    Public Const TRACK_IMG_LEFT = 210
    ' 配置位置y
    Public Const TRACK_IMG_TOP = 100
    ' 配置サイズx
    Public Const TRACK_IMG_WIDTH = 225
    ' 配置サイズy
    Public Const TRACK_IMG_HEIGHT = 225
    

' 初期選択セル
    Public Const INIT_SELECT_REGIST_DATA = "C2"
    Public Const INIT_SELECT_GRAPH = "B1"
    Public Const INIT_SELECT_DATA = "A1"
    Public Const INIT_SELECT_SETTINGS = "A1"

' セル座標 - RegistData
    Public Const REGIST_ROW_TIER = 2
    Public Const REGIST_COL_TIER = 3
    Public Const REGIST_ROW_FORMAT = 3
    Public Const REGIST_COL_FORMAT = 3
    Public Const REGIST_ROW_HEADER = 4
    Public Const REGIST_COL_TRACK_NAME = 2
    Public Const REGIST_COL_START_RANK = 3
    Public Const REGIST_COL_RANK = 4
    Public Const REGIST_COL_REMARK = 5
    Public Const REGIST_ROW_KNOWLEDGE = 18
    Public Const REGIST_COL_KNOWLEDGE = 3

' セル座標 - STORAGE
    Public Const STORAGE_COL_TRACK_NAME = 1
    Public Const STORAGE_ROW_TRACK_NAME = 1
    Public Const STORAGE_COL_TRACK_KEY = 2
    Public Const STORAGE_ROW_TRACK_KEY = 2
    Public Const STORAGE_COL_LANGUAGE_NAME = 3
    Public Const STORAGE_Row_LANGUAGE_NAME = 1
    Public Const STORAGE_COL_LANGUAGE_KEY = 4
    Public Const STORAGE_Row_LANGUAGE_KEY = 1
    Public Const STORAGE_COL_TIER_NAME = 5
    Public Const STORAGE_ROW_TIER_NAME = 1
    Public Const STORAGE_COL_FORMAT_NAME = 6
    Public Const STORAGE_ROW_FORMAT_NAME = 1
    Public Const STORAGE_COL_POINT_FLG = 7
    Public Const STORAGE_ROW_POINT_FLG = 1
    Public Const STORAGE_COL_RANK_FLG = 8
    Public Const STORAGE_ROW_RANK_FLG = 1
    Public Const STORAGE_COL_TIMES_FLG = 9
    Public Const STORAGE_ROW_TIMES_FLG = 1
    

' セル座標 - SETTINGS
    Public Const SETTINGS_COL_KEY = 1
    Public Const SETTINGS_COL_VALUE = 2
    Public Const SETTINGS_COL_DISPLAY = 3
    Public Const SETTINGS_ROW_LANGUAGE = 3
    Public Const SETTINGS_ROW_RACE_NUM = 4
    Public Const SETTINGS_ROW_SHOWING_MAP = 5

' セル座標 - Data
    Public Const DATA_ROW_HEADER = 1
    Public Const DATA_COL_REGIST_KEY = 1
    Public Const DATA_COL_DATE = 2
    Public Const DATA_COL_TIER = 3
    Public Const DATA_COL_FORMAT = 4
    Public Const DATA_COL_TRACK_KEY = 5
    Public Const DATA_COL_TRACK_NAME_JP = 6
    Public Const DATA_COL_TRACK_NAME_EN = 7
    Public Const DATA_COL_STARTING_RANK = 8
    Public Const DATA_COL_RANK = 9
    Public Const DATA_COL_POINT = 10
    Public Const DATA_COL_REMARK = 11
    




