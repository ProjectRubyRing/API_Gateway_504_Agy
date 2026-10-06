import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def create_workbook():
    wb = openpyxl.Workbook()
    # remove default sheet later or rename it
    ws_default = wb.active
    ws_default.title = "01_はじめにと小学生向け解説"

    # Define common styles
    FONT_NAME = "Meiryo UI"
    
    # Fonts
    font_title = Font(name=FONT_NAME, size=16, bold=True, color="FFFFFF")
    font_subtitle = Font(name=FONT_NAME, size=11, italic=True, color="D9D9D9")
    font_section = Font(name=FONT_NAME, size=13, bold=True, color="FFFFFF")
    font_subsection = Font(name=FONT_NAME, size=11, bold=True, color="262626")
    font_th = Font(name=FONT_NAME, size=10, bold=True, color="FFFFFF")
    font_tb = Font(name=FONT_NAME, size=10, color="1A1A1A")
    font_tb_bold = Font(name=FONT_NAME, size=10, bold=True, color="1A1A1A")
    font_code = Font(name="Consolas", size=9, color="262626")
    font_note = Font(name=FONT_NAME, size=9, italic=True, color="595959")
    
    # Fills (Monotone Palette)
    fill_title = PatternFill(start_color="1A1A1A", end_color="1A1A1A", fill_type="solid")      # Almost black
    fill_section = PatternFill(start_color="333333", end_color="333333", fill_type="solid")    # Dark gray
    fill_th = PatternFill(start_color="4D4D4D", end_color="4D4D4D", fill_type="solid")         # Charcoal gray
    fill_th_sub = PatternFill(start_color="737373", end_color="737373", fill_type="solid")     # Medium gray
    fill_zebra = PatternFill(start_color="F7F7F7", end_color="F7F7F7", fill_type="solid")      # Very light gray
    fill_white = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    fill_highlight = PatternFill(start_color="E6E6E6", end_color="E6E6E6", fill_type="solid")  # Accent light gray
    fill_callout = PatternFill(start_color="EEEEEE", end_color="EEEEEE", fill_type="solid")    # Note background

    # Borders
    border_thin = Border(
        left=Side(style='thin', color='CCCCCC'),
        right=Side(style='thin', color='CCCCCC'),
        top=Side(style='thin', color='CCCCCC'),
        bottom=Side(style='thin', color='CCCCCC')
    )
    border_th = Border(
        left=Side(style='thin', color='666666'),
        right=Side(style='thin', color='666666'),
        top=Side(style='medium', color='1A1A1A'),
        bottom=Side(style='medium', color='1A1A1A')
    )
    border_box = Border(
        left=Side(style='medium', color='4D4D4D'),
        right=Side(style='medium', color='4D4D4D'),
        top=Side(style='medium', color='4D4D4D'),
        bottom=Side(style='medium', color='4D4D4D')
    )

    # Alignments
    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    align_top_left = Alignment(horizontal="left", vertical="top", wrap_text=True)
    align_top_center = Alignment(horizontal="center", vertical="top", wrap_text=True)
    align_right = Alignment(horizontal="right", vertical="center", wrap_text=True)

    # -------------------------------------------------------------
    # SHEET 1: 01_はじめにと小学生向け解説
    # -------------------------------------------------------------
    ws1 = ws_default
    ws1.views.sheetView[0].showGridLines = True

    # Title Block
    ws1.merge_cells("A1:G1")
    ws1["A1"] = "【超徹底解剖】API Gateway 504 即時返却・バックエンドALB 200 OK の謎を解く"
    ws1["A1"].font = font_title
    ws1["A1"].fill = fill_title
    ws1["A1"].alignment = align_center
    ws1.row_dimensions[1].height = 40

    ws1.merge_cells("A2:G2")
    ws1["A2"] = "〜 レストランのすれ違い事件でわかる！VPC Link v2 × REST API × ALB における即時504タイムアウトの全貌 〜"
    ws1["A2"].font = font_subtitle
    ws1["A2"].fill = fill_title
    ws1["A2"].alignment = align_center
    ws1.row_dimensions[2].height = 24

    ws1.row_dimensions[3].height = 15

    # Section 1
    ws1.merge_cells("A4:G4")
    ws1["A4"] = "第1章：何が起きているのか？（小学生にもわかる超直感たとえ話）"
    ws1["A4"].font = font_section
    ws1["A4"].fill = fill_section
    ws1["A4"].alignment = align_left
    ws1.row_dimensions[4].height = 28

    story_text = (
        "【あるレストランの不思議な事件】\n"
        "お客さんがレストランに来て『ハンバーグ定食をください！』と注文しました。\n"
        "すると、注文した【0.1秒後（まばたきする間）】に、受付のロボット店員が突然、\n"
        "『時間切れです！注文はキャンセルされました！（504 Gateway Timeout）』と叫びました。\n\n"
        "お客さんはビックリして『えっ！？いま頼んだばかりなのに時間切れ！？』と怒りました。\n"
        "そこで厨房（裏口の安全な部屋）をのぞいてみると…なんと料理長（ALB）とコックさん（バックエンド）は\n"
        "『よし！絶品ハンバーグ完成！大成功（200 OK）！』と満面の笑みで料理をお盆に乗せていたのです！\n\n"
        "厨房では『大成功（200 OK）』なのに、なぜお客さんには『一瞬で時間切れ（504）』と突き返されたのでしょうか？"
    )
    ws1.merge_cells("A5:G5")
    ws1["A5"] = story_text
    ws1["A5"].font = font_tb
    ws1["A5"].fill = fill_callout
    ws1["A5"].alignment = align_top_left
    ws1.row_dimensions[5].height = 140

    ws1.row_dimensions[6].height = 15

    # Characters Table
    ws1.merge_cells("A7:G7")
    ws1["A7"] = "■ クラウド世界の登場人物とレストランの役割対応表"
    ws1["A7"].font = font_subsection
    ws1.row_dimensions[7].height = 24

    headers1 = ["役割", "IT用語（クラウド側）", "レストランのたとえ", "性格・特徴・お仕事", "今回の事件での状態"]
    col_widths1 = [12, 24, 20, 42, 35]
    
    # Write headers
    row_idx = 8
    ws1.row_dimensions[row_idx].height = 26
    ws1.merge_cells(f"D{row_idx}:E{row_idx}")
    ws1.merge_cells(f"F{row_idx}:G{row_idx}")
    
    ws1[f"A{row_idx}"] = "役割"
    ws1[f"B{row_idx}"] = "IT用語（クラウド側）"
    ws1[f"C{row_idx}"] = "レストランのたとえ"
    ws1[f"D{row_idx}"] = "性格・特徴・お仕事"
    ws1[f"F{row_idx}"] = "今回の事件での状態"
    
    for col_c in ["A", "B", "C", "D", "E", "F", "G"]:
        ws1[f"{col_c}{row_idx}"].font = font_th
        ws1[f"{col_c}{row_idx}"].fill = fill_th
        ws1[f"{col_c}{row_idx}"].alignment = align_center
        ws1[f"{col_c}{row_idx}"].border = border_th

    char_data = [
        ("クライアント", "スマホ / Webブラウザ / 外部API", "お店に来たお客さん", "画面をポチッと押してデータを要求する。結果が戻るのを待っている。", "リクエスト直後に即座に『504』エラーを突き返されて困惑している。"),
        ("受付窓口", "API Gateway (REST API)", "せっかちな受付ロボット店員", "玄関でお客さんの注文を受け、チェック（認証や制限）して厨房へ送る。自分の時計で時間を計り、約束の時間を1ミリ秒でも超えると『注文キャンセル！』と叫ぶ。", "『時間切れ！』と叫んでお客さんに504を返却。しかし厨房へ注文は送ってしまっていた。"),
        ("専用通路", "VPC Link v2 (Hyperplane ENI)", "厨房へ通じる特急シューター", "一般人が入れない秘密の厨房（VPC）へ、カプセルを高速で飛ばす専用トンネル。双方向で通信できる。", "注文カプセルを厨房へ送り届ける。しかし受付ロボットが切断したためパケットの行き場がなくなる。"),
        ("厨房の司令塔", "ALB (Application Load Balancer)", "厨房の料理長（チーフ）", "届いた伝票を見て『この料理はあっちのコック、あの料理はこっちのコック』と振り分け、出来上がりを点検して『200点（合格）！』と記録する。", "コックから出来上がった料理を受け取り『大成功！（200 OK）』と伝票にハンコを押した。"),
        ("料理人", "バックエンド (ECS / EC2 / Lambda)", "実際にフライパンを振るコック", "材料を炒め、データベースからデータを取り出して美味しいデータ料理を完成させる。", "指示通りに一生懸命調理して、見事に完成させて料理長へ渡した。何一つ悪くない。")
    ]

    row_idx = 9
    for item in char_data:
        ws1.row_dimensions[row_idx].height = 42
        ws1.merge_cells(f"D{row_idx}:E{row_idx}")
        ws1.merge_cells(f"F{row_idx}:G{row_idx}")
        
        ws1[f"A{row_idx}"] = item[0]
        ws1[f"B{row_idx}"] = item[1]
        ws1[f"C{row_idx}"] = item[2]
        ws1[f"D{row_idx}"] = item[3]
        ws1[f"F{row_idx}"] = item[4]
        
        fill_curr = fill_zebra if row_idx % 2 == 1 else fill_white
        for col_c in ["A", "B", "C", "D", "E", "F", "G"]:
            cell = ws1[f"{col_c}{row_idx}"]
            cell.font = font_tb
            cell.fill = fill_curr
            cell.border = border_thin
            if col_c in ["A", "B", "C"]:
                cell.alignment = align_center
            else:
                cell.alignment = align_left
        row_idx += 1

    row_idx += 1
    # Section Summary Box
    ws1.merge_cells(f"A{row_idx}:G{row_idx}")
    ws1[f"A{row_idx}"] = "■ なぜこんな『すれ違い』が起きるの？ 3つの大きな理由（直感まとめ）"
    ws1[f"A{row_idx}"].font = font_subsection
    ws1.row_dimensions[row_idx].height = 24
    row_idx += 1

    summary_bullets = [
        "① ロボット店員の時計が壊れていた（設定ミス）: 本当は『29秒待つ』つもりが、単位を間違えて『0.05秒（50ミリ秒）』でセットされていた！コックは0.1秒で作ったのにロボットは0.05秒で勝手に怒ってキャンセルした。",
        "② 料理長（ALB）の日記に書かれた『大成功（200）』は、実は別の料理（ヘルスチェック等）だった！: 料理長が定期的に自分で味見して『うん、200点！』と書いていた日記を見て、エンジニアが勘違いした。",
        "③ 伝票のやりとり中に糸電話の糸がプツンと切れた（TCP切断・プロトコル異常）: 料理長は完成させたつもりだったが、ロボット店員との通信路が途中で壊れ、ロボット店員は『届かないから時間切れ！』と判断した。"
    ]

    for b in summary_bullets:
        ws1.merge_cells(f"A{row_idx}:G{row_idx}")
        ws1[f"A{row_idx}"] = b
        ws1[f"A{row_idx}"].font = font_tb
        ws1[f"A{row_idx}"].fill = fill_white
        ws1[f"A{row_idx}"].alignment = align_top_left
        ws1[f"A{row_idx}"].border = border_thin
        ws1.row_dimensions[row_idx].height = 36
        row_idx += 1

    # Set column widths
    ws1.column_dimensions["A"].width = 16
    ws1.column_dimensions["B"].width = 28
    ws1.column_dimensions["C"].width = 22
    ws1.column_dimensions["D"].width = 28
    ws1.column_dimensions["E"].width = 28
    ws1.column_dimensions["F"].width = 28
    ws1.column_dimensions["G"].width = 28

    # -------------------------------------------------------------
    # SHEET 2: 02_歴史とVPC_Link_v2誕生経緯
    # -------------------------------------------------------------
    ws2 = wb.create_sheet(title="02_歴史とVPC_Link_v2誕生経緯")
    ws2.views.sheetView[0].showGridLines = True

    # Title Block
    ws2.merge_cells("A1:G1")
    ws2["A1"] = "第2章：API Gatewayとプライベート統合の歴史とVPC Link v2誕生の経緯"
    ws2["A1"].font = font_title
    ws2["A1"].fill = fill_title
    ws2["A1"].alignment = align_center
    ws2.row_dimensions[1].height = 40

    ws2.merge_cells("A2:G2")
    ws2["A2"] = "〜 なぜVPC Link v2が必要だったのか？ 悪名高き『NLB-ALBサンドイッチ』からALB直結解禁までの全歴史 〜"
    ws2["A2"].font = font_subtitle
    ws2["A2"].fill = fill_title
    ws2["A2"].alignment = align_center
    ws2.row_dimensions[2].height = 24

    ws2.row_dimensions[3].height = 15

    # History Table
    row_idx = 4
    ws2.merge_cells(f"A{row_idx}:G{row_idx}")
    ws2[f"A{row_idx}"] = "■ API Gateway プライベート通信の進化史（2015年〜現在）"
    ws2[f"A{row_idx}"].font = font_section
    ws2[f"A{row_idx}"].fill = fill_section
    ws2[f"A{row_idx}"].alignment = align_left
    ws2.row_dimensions[row_idx].height = 28
    row_idx += 1

    hist_headers = ["年代・フェーズ", "アーキテクチャ概要", "技術的仕組み", "メリット", "最大の課題・エンジニアの悲鳴"]
    ws2.row_dimensions[row_idx].height = 26
    ws2.merge_cells(f"C{row_idx}:D{row_idx}")
    ws2.merge_cells(f"F{row_idx}:G{row_idx}")
    
    ws2[f"A{row_idx}"] = "年代・フェーズ"
    ws2[f"B{row_idx}"] = "アーキテクチャ構成"
    ws2[f"C{row_idx}"] = "技術的仕組み"
    ws2[f"E{row_idx}"] = "当時のメリット"
    ws2[f"F{row_idx}"] = "最大の課題・エンジニアの悲鳴"

    for col_c in ["A", "B", "C", "D", "E", "F", "G"]:
        ws2[f"{col_c}{row_idx}"].font = font_th
        ws2[f"{col_c}{row_idx}"].fill = fill_th
        ws2[f"{col_c}{row_idx}"].alignment = align_center
        ws2[f"{col_c}{row_idx}"].border = border_th
    row_idx += 1

    history_data = [
        (
            "時代①：2015年〜\n【黎明期】\nREST API誕生",
            "API GW (Public)\n      ↓ (Internet)\nInternet Gateway\n      ↓\nPublic ALB / EC2",
            "VPC統合機能が存在せず、API GatewayからVPC内のサーバーを呼ぶにはパブリックIP経由で通信するしかなかった。",
            "サーバーレス（Lambda）やパブリックAPIの共通基盤として急速に普及。",
            "社内専用APIやDB接続アプリをVPC内に隠蔽できず、セキュリティ要件が厳しいエンタープライズ企業で採用できなかった。"
        ),
        (
            "時代②：2017年〜\n【暗黒期】\nVPC Link v1と\nNLB-ALBサンドイッチ",
            "API GW (REST)\n      ↓\nVPC Link v1 (PrivateLink)\n      ↓\nNLB (L4)\n      ↓\nALB (L7) → ECS/EC2",
            "AWS PrivateLink技術に基づく『VPC Link v1』が登場。しかし仕様上【NLBにしか繋げない】制限があった。\nWebで必須のALBを使うため、NLBの後ろにALBを置く二重構成が必須となった。",
            "VPC内の完全プライベートなリソースへ安全に通信可能となった。",
            "【エンジニア全員が絶望したNLB-ALBサンドイッチ】\n・ロードバランサーが2台必要でコスト2倍！\n・無駄な1ホップでレイテンシ悪化\n・NLBにALBのIPを登録する同期Lambdaが必要など運用が地獄！"
        ),
        (
            "時代③：2019年末〜\n【分断期】\nHTTP API & VPC Link v2\nREST APIの取り残し",
            "【HTTP APIの場合】\nHTTP API → VPC Link v2 → ALB\n\n【REST APIの場合】\nREST API → VPC Link v1 → NLB → ALB",
            "軽量な『HTTP API』と共にAWS Hyperplane技術を活用した『VPC Link v2』が登場。ALBへ直接接続可能となった。\nしかし！VPC Link v2はHTTP API専用で、REST APIには提供されなかった。",
            "HTTP APIではNLBが不要になり、コスト大幅減・構成が超シンプルに。",
            "既存のREST APIはAPIキー管理、WAF、詳細マッピング等のためHTTP APIへ移行できず、NLB-ALBサンドイッチを使い続けなければならなかった（長い我慢の時代）。"
        ),
        (
            "時代④：2024年11月/2025年〜\n【最新・解放期】\nREST APIが\nVPC Link v2に対応！",
            "API GW (REST API)\n      ↓\nVPC Link v2 (Hyperplane)\n      ↓\nALB (L7 Direct) → ECS/EC2",
            "AWS公式アップデートにより、REST APIでもVPC Link v2（ALB直接統合およびCloud Map統合）が正式サポート！\n中間NLBを完全撤廃できるようになった。",
            "・中間NLBのコスト完全撤廃！\n・レイテンシ大幅削減！\n・IaC定義が激的にシンプル化！\n・ALBのL7機能（パスルーティング等）をダイレクト活用可能。",
            "【新たなトラブルの発生】\nNLBを抜いたことで、API Gatewayの厳密なプロキシ仕様とALBの挙動が直面し、『ALBは200なのに即時504エラー』などの新特有の課題が噴出！"
        )
    ]

    for item in history_data:
        ws2.row_dimensions[row_idx].height = 80
        ws2.merge_cells(f"C{row_idx}:D{row_idx}")
        ws2.merge_cells(f"F{row_idx}:G{row_idx}")
        
        ws2[f"A{row_idx}"] = item[0]
        ws2[f"B{row_idx}"] = item[1]
        ws2[f"C{row_idx}"] = item[2]
        ws2[f"E{row_idx}"] = item[3]
        ws2[f"F{row_idx}"] = item[4]
        
        fill_curr = fill_zebra if (row_idx % 2 == 1) else fill_white
        for col_c in ["A", "B", "C", "D", "E", "F", "G"]:
            cell = ws2[f"{col_c}{row_idx}"]
            cell.font = font_tb
            cell.fill = fill_curr
            cell.border = border_thin
            if col_c in ["A", "B"]:
                cell.alignment = align_center
            else:
                cell.alignment = align_left
        row_idx += 1

    row_idx += 1
    # Deep dive on VPC Link v1 vs v2
    ws2.merge_cells(f"A{row_idx}:G{row_idx}")
    ws2[f"A{row_idx}"] = "■ 徹底比較：VPC Link v1 と VPC Link v2 のアーキテクチャ的相違点"
    ws2[f"A{row_idx}"].font = font_section
    ws2[f"A{row_idx}"].fill = fill_section
    ws2[f"A{row_idx}"].alignment = align_left
    ws2.row_dimensions[row_idx].height = 28
    row_idx += 1

    v1_v2_headers = ["比較項目", "従来の VPC Link v1", "最新の VPC Link v2", "アーキテクチャへの影響・注意点"]
    ws2.row_dimensions[row_idx].height = 26
    ws2.merge_cells(f"D{row_idx}:G{row_idx}")
    
    ws2[f"A{row_idx}"] = "比較項目"
    ws2[f"B{row_idx}"] = "従来の VPC Link v1"
    ws2[f"C{row_idx}"] = "最新の VPC Link v2"
    ws2[f"D{row_idx}"] = "アーキテクチャへの影響・注意点"

    for col_c in ["A", "B", "C", "D", "E", "F", "G"]:
        ws2[f"{col_c}{row_idx}"].font = font_th
        ws2[f"{col_c}{row_idx}"].fill = fill_th
        ws2[f"{col_c}{row_idx}"].alignment = align_center
        ws2[f"{col_c}{row_idx}"].border = border_th
    row_idx += 1

    v1_v2_rows = [
        ("下層ネットワーク技術", "AWS PrivateLink", "AWS Hyperplane (NAT/クロスアカウント分散技術)", "HyperplaneはNAT GatewayやNLBの裏側でも動く超高スループット基盤。"),
        ("直接接続可能なターゲット", "NLB (Network Load Balancer) のみ", "ALB (Application Load Balancer)、NLB、Cloud Map", "ALBへ直接通信できるため、中間の不要なNLBを完全に削除可能。"),
        ("セキュリティグループ (SG)", "VPC Link自体にSGは付与できない (NLB側で制御)", "VPC Link v2 自体にセキュリティグループを付与可能", "API Gateway側のENIとALB側のSGで双方向の通信許可設計が極めて重要。"),
        ("対応APIタイプ", "REST API のみ", "HTTP API および REST API (最新アップデート)", "REST APIでも最新のVPC Link v2が共通で利用できるようになった。"),
        ("ロードバランサーとの関係", "1つのVPC Linkにつき1つのNLB", "1つのVPC Link v2から複数のALB/NLBへ接続可能 (1対多)", "1つのVPC Linkリソースを複数のマイクロサービスALBで相乗り・再利用可能。")
    ]

    for item in v1_v2_rows:
        ws2.row_dimensions[row_idx].height = 36
        ws2.merge_cells(f"D{row_idx}:G{row_idx}")
        ws2[f"A{row_idx}"] = item[0]
        ws2[f"B{row_idx}"] = item[1]
        ws2[f"C{row_idx}"] = item[2]
        ws2[f"D{row_idx}"] = item[3]
        
        fill_curr = fill_zebra if (row_idx % 2 == 1) else fill_white
        for col_c in ["A", "B", "C", "D", "E", "F", "G"]:
            cell = ws2[f"{col_c}{row_idx}"]
            cell.font = font_tb
            cell.fill = fill_curr
            cell.border = border_thin
            if col_c in ["A", "B", "C"]:
                cell.alignment = align_center
            else:
                cell.alignment = align_left
        row_idx += 1

    ws2.column_dimensions["A"].width = 22
    ws2.column_dimensions["B"].width = 28
    ws2.column_dimensions["C"].width = 30
    ws2.column_dimensions["D"].width = 25
    ws2.column_dimensions["E"].width = 25
    ws2.column_dimensions["F"].width = 25
    ws2.column_dimensions["G"].width = 25

    # -------------------------------------------------------------
    # SHEET 3: 03_アーキテクチャと動作原理
    # -------------------------------------------------------------
    ws3 = wb.create_sheet(title="03_アーキテクチャと動作原理")
    ws3.views.sheetView[0].showGridLines = True

    # Title Block
    ws3.merge_cells("A1:G1")
    ws3["A1"] = "第3章：動作原理と通信シーケンスの深堀り"
    ws3["A1"].font = font_title
    ws3["A1"].fill = fill_title
    ws3["A1"].alignment = align_center
    ws3.row_dimensions[1].height = 40

    ws3.merge_cells("A2:G2")
    ws3["A2"] = "〜 API Gateway REST API・Hyperplane ENI・ALBのパケット送受信とタイマー機構 〜"
    ws3["A2"].font = font_subtitle
    ws3["A2"].fill = fill_title
    ws3["A2"].alignment = align_center
    ws3.row_dimensions[2].height = 24

    ws3.row_dimensions[3].height = 15

    row_idx = 4
    ws3.merge_cells(f"A{row_idx}:G{row_idx}")
    ws3[f"A{row_idx}"] = "■ パケット伝送フローと3つの重要な動作原理"
    ws3[f"A{row_idx}"].font = font_section
    ws3[f"A{row_idx}"].fill = fill_section
    ws3[f"A{row_idx}"].alignment = align_left
    ws3.row_dimensions[row_idx].height = 28
    row_idx += 1

    principles = [
        (
            "① API Gateway REST API の『完全バッファリング型プロキシ』仕様",
            "HTTP APIや最新のリバースプロキシと異なり、API Gateway REST API はレスポンスストリーミング（Chunked転送の逐次返却）をサポートしません。\n"
            "バックエンドから返ってくるHTTPレスポンスの『すべてのバイト（ヘッダー＋全ボディ）』を受信し、API Gateway内部のメモリにすべて溜め込んでからでないと、クライアントにレスポンスを返せません。\n"
            "⇒ 影響：途中でボディが途切れたり、1バイトでも不整合があると、全体が即時破棄され502/504エラーになります。"
        ),
        (
            "② API Gateway の『統合タイムアウト（Integration Timeout）』タイマー機構",
            "API Gateway REST API のメソッド統合には『timeoutInMillis』という設定値があります（デフォルトは29,000ms = 29秒、設定可能範囲は50ms〜29,000ms）。\n"
            "API GatewayがVPC Link経由でALBへリクエストを投げた瞬間から、内部の独立した高精度タイマーがカウントダウンを開始します。\n"
            "タイマーが満了すると、ALBが処理中であろうがなかろうが、API Gatewayは即座に通信を中断し、クライアントへ『HTTP 504 Gateway Timeout』を返却します。\n"
            "⇒ 影響：設定値が例えば『50ms』や『100ms』になっていると、人間からは『即時返却』に見える超高速504が発生します。"
        ),
        (
            "③ ALB (Application Load Balancer) の『ステータスコード決定とログ出力タイミング』",
            "ALBはクライアント側とバックエンド側の2つのTCPコネクションを仲介するリバースプロキシです。\n"
            "ALBがアクセスログに記録する『elb_status_code』と『target_status_code』は、ターゲット（コンテナ等）からHTTPヘッダー『HTTP/1.1 200 OK』を受け取った時点で『200』と確定します。\n"
            "ALBが200をAPI Gatewayに送出しても、API Gateway側がすでにタイムアウトでソケットを閉じていた場合、ALB側は『自分は200を返して仕事を終えた』と記録するため、ALBログには『200』が残り、クライアントには『504』が届くという決定的なズレが発生します。"
        )
    ]

    for p in principles:
        ws3.merge_cells(f"A{row_idx}:G{row_idx}")
        ws3[f"A{row_idx}"] = p[0]
        ws3[f"A{row_idx}"].font = font_subsection
        ws3[f"A{row_idx}"].fill = fill_highlight
        ws3[f"A{row_idx}"].alignment = align_left
        ws3[f"A{row_idx}"].border = border_thin
        ws3.row_dimensions[row_idx].height = 24
        row_idx += 1

        ws3.merge_cells(f"A{row_idx}:G{row_idx}")
        ws3[f"A{row_idx}"] = p[1]
        ws3[f"A{row_idx}"].font = font_tb
        ws3[f"A{row_idx}"].fill = fill_white
        ws3[f"A{row_idx}"].alignment = align_top_left
        ws3[f"A{row_idx}"].border = border_thin
        ws3.row_dimensions[row_idx].height = 64
        row_idx += 1

    row_idx += 1
    # Sequence Diagram in ASCII/Text
    ws3.merge_cells(f"A{row_idx}:G{row_idx}")
    ws3[f"A{row_idx}"] = "■ 正常通信時 vs 即時504発生時の動作イメージ（時系列比較）"
    ws3[f"A{row_idx}"].font = font_section
    ws3[f"A{row_idx}"].fill = fill_section
    ws3[f"A{row_idx}"].alignment = align_left
    ws3.row_dimensions[row_idx].height = 28
    row_idx += 1

    seq_headers = ["区分", "クライアント", "API Gateway (REST)", "VPC Link v2 (Hyperplane)", "ALB (L7)", "バックエンド (ECS)", "結果ログ"]
    ws3.row_dimensions[row_idx].height = 26
    for idx_c, h in enumerate(seq_headers):
        col_letter = get_column_letter(idx_c + 1)
        ws3[f"{col_letter}{row_idx}"] = h
        ws3[f"{col_letter}{row_idx}"].font = font_th
        ws3[f"{col_letter}{row_idx}"].fill = fill_th
        ws3[f"{col_letter}{row_idx}"].alignment = align_center
        ws3[f"{col_letter}{row_idx}"].border = border_th
    row_idx += 1

    seq_data = [
        ("正常時", "リクエスト送信", "リクエスト受信\nタイマー開始 (29秒)", "パケット転送", "リクエスト受信\nターゲットへ転送", "処理実行 (100ms)\n200 OK 返却", "ALB: 200 OK\nAPI GW: 200 OK\nClient: 200 OK"),
        ("異常時①\n統合TO極小", "リクエスト送信", "リクエスト転送後…\nわずか50msでタイマー切れ！\n→ 504即時返却！", "通信中断 (RST/FIN)", "遅れて200 OKを受信\nAPI GWへ送出試行\n(届かず)", "通常通り100msで完了\n200 OK返却", "ALB: 200 OK\nAPI GW: 504 Timeout\nClient: 504 (所要50ms)"),
        ("異常時②\nKeep-Alive競合", "リクエスト送信", "既存TCPプールから選択\nリクエスト送信\nALBからRST受信で即死", "パケット転送\n(FINと衝突)", "ちょうど60秒アイドルで\nTCP FIN送信済み\n→ 新データにRST応答", "前リクエストは正常終了\nまたは未到達", "ALB: 200 (直前ログ)\nAPI GW: 504 / 502\nClient: 504 即時"),
        ("異常時③\nChunk途中死", "リクエスト送信", "リクエスト転送\nレスポンス受信開始\n途中で通信途絶を検知", "パケット転送", "ヘッダー受信で200確定\nボディ中継中にターゲット死亡", "ヘッダー200送出後\nOOM等で突然プロセス死", "ALB: 200 (ヘッダー済)\nAPI GW: 504 / 502\nClient: 504 即時")
    ]

    for item in seq_data:
        ws3.row_dimensions[row_idx].height = 58
        for idx_c, val in enumerate(item):
            col_letter = get_column_letter(idx_c + 1)
            cell = ws3[f"{col_letter}{row_idx}"]
            cell.value = val
            cell.font = font_tb
            cell.fill = fill_zebra if (row_idx % 2 == 1) else fill_white
            cell.border = border_thin
            if idx_c in [0, 6]:
                cell.alignment = align_center
            else:
                cell.alignment = align_left
        row_idx += 1

    ws3.column_dimensions["A"].width = 16
    ws3.column_dimensions["B"].width = 22
    ws3.column_dimensions["C"].width = 26
    ws3.column_dimensions["D"].width = 24
    ws3.column_dimensions["E"].width = 26
    ws3.column_dimensions["F"].width = 24
    ws3.column_dimensions["G"].width = 24

    # -------------------------------------------------------------
    # SHEET 4: 04_504即時返却の全ケース一覧
    # -------------------------------------------------------------
    ws4 = wb.create_sheet(title="04_504即時返却の全ケース一覧")
    ws4.views.sheetView[0].showGridLines = True

    # Title Block
    ws4.merge_cells("A1:G1")
    ws4["A1"] = "第4章：504即時返却・バックエンドALB 200 OK が発生する全ケース一覧マトリクス"
    ws4["A1"].font = font_title
    ws4["A1"].fill = fill_title
    ws4["A1"].alignment = align_center
    ws4.row_dimensions[1].height = 40

    ws4.merge_cells("A2:G2")
    ws4["A2"] = "〜 条件：REST API × VPC Link v2 × ALB で504が即時返り、ALBが200を返す全パターンの比較網羅 〜"
    ws4["A2"].font = font_subtitle
    ws4["A2"].fill = fill_title
    ws4["A2"].alignment = align_center
    ws4.row_dimensions[2].height = 24

    ws4.row_dimensions[3].height = 15

    row_idx = 4
    ws4.merge_cells(f"A{row_idx}:G{row_idx}")
    ws4[f"A{row_idx}"] = "■ 原因ケース比較・分類マトリクス"
    ws4[f"A{row_idx}"].font = font_section
    ws4[f"A{row_idx}"].fill = fill_section
    ws4[f"A{row_idx}"].alignment = align_left
    ws4.row_dimensions[row_idx].height = 28
    row_idx += 1

    case_headers = ["No", "ケース分類・名称", "504が『即時』返却される根本メカニズム", "ALBが『200』で返却/記録する理由", "発生頻度", "調査の着眼点 (決め手)", "即効の対策・解決策"]
    col_widths4 = [6, 26, 38, 38, 12, 34, 30]
    
    ws4.row_dimensions[row_idx].height = 28
    for idx_c, h in enumerate(case_headers):
        col_letter = get_column_letter(idx_c + 1)
        ws4[f"{col_letter}{row_idx}"] = h
        ws4[f"{col_letter}{row_idx}"].font = font_th
        ws4[f"{col_letter}{row_idx}"].fill = fill_th
        ws4[f"{col_letter}{row_idx}"].alignment = align_center
        ws4[f"{col_letter}{row_idx}"].border = border_th
    row_idx += 1

    all_cases_matrix = [
        (
            "1",
            "【設定の単位ミス】\nAPI Gatewayの統合タイムアウトが極小値（50ms〜数百ms）",
            "TerraformやCloudFormationで『timeout_in_millis = 29』等と秒とミリ秒を混同して設定（API GWの最小値50msが適用）。50msや100msでAPI GWがタイムアウトし、クライアントへ504を即時返却。",
            "ALBとバックエンドは通常通り処理（例: 150msで正常完了）。ターゲットが200を返し、ALBはアクセスログに『elb_status_code=200, target_status_code=200』を記録。",
            "★★★★★\n(最も多い)",
            "API Gateway実行ログの経過時間（約50ms〜100msでMethod completed: 504）。ALBログのtarget_processing_timeがその値を超過。",
            "IaCやコンソールで timeoutInMillis を 29000 (29秒) や適切な値（例: 5000ms）に修正してステージを再デプロイ。"
        ),
        (
            "2",
            "【調査の錯覚・別人事件】\nALBログの200は別リクエスト (ヘルスチェックや別通信)",
            "実際のリクエストはVPC Link v2のセキュリティグループやNACL、リスナー設定不備でALBに届かず、API GWが接続拒否/失敗を検知して即時504を返却。",
            "調査時に見たALBログの『200』は、ターゲットグループの定期ヘルスチェック（GET /health）や別ユーザーのリクエスト、またはUTCとJSTの9時間ズレによる過去ログだった！",
            "★★★★☆\n(極めて多い)",
            "API Gatewayの『x-amzn-trace-id』とALBログの『trace_id』を完全一致検索すると、実際のリクエストがALBに1件も存在しない。",
            "X-Ray Trace IDでログを確実に1対1突合。VPC Link v2とALBのSecurity Group（ポート通信許可）を修正。"
        ),
        (
            "3",
            "【TCPキープアライブ競合】\nKeep-Alive再利用とALBアイドルタイムアウトの衝突",
            "API Gatewayがプーリングしていた既存TCPコネクションを再利用した瞬間、ALB側がアイドルタイムアウト（60秒）によりTCP FINを切断送信。すれ違いでTCP RSTが発生し、API GWが即時エラー検知。",
            "ALBがリクエストを一部受け取ってバックエンドへ転送し、バックエンドが200で終えた直後にRST切断されたか、直前リクエストのログと誤認。",
            "★★★☆☆\n(高負荷・稀)",
            "API GW実行ログに『Connection reset by peer』等の切断エラー。発生間隔がアイドルタイムアウト値（60秒）付近に集中。",
            "ALBのアイドルタイムアウトをAPI Gatewayのキープアライブ保持時間より長め（例: 65秒以上）に設定、またはバックエンドKeep-Alive調整。"
        ),
        (
            "4",
            "【HTTPプロトコル違反】\nChunked Transfer Encoding の途絶・未完結",
            "バックエンドが『Transfer-Encoding: chunked』で応答中、途中でコンテナがOOM Killer等で突然死。終端チャンク（0\\r\\n\\r\\n）が届かずソケット切断され、バッファリング中のAPI GWが即時504/502返却。",
            "ターゲットがステータスヘッダー『HTTP/1.1 200 OK』をALBに送った時点で、ALBはレスポンスコードを200と判定してログ出力。",
            "★★★☆☆\n(アプリ異常)",
            "ALBアクセスログの sent_bytes / received_bytes が異常に小さい。バックエンドECSタスクのStoppedReasonに『OutOfMemory』等がある。",
            "バックエンドのメモリ不足解消、例外処理の修正。REST APIの前に完全なContent-Lengthで返却する設計にする。"
        ),
        (
            "5",
            "【Content-Length不整合】\n宣言サイズと実送信サイズの不一致によるソケット強制終了",
            "バックエンドが『Content-Length: 5000』と返したのに、実際には200バイト送って接続を終了。API GWは残りデータを待つがソケット切断され、即座にゲートウェイ異常（504/502）と判定。",
            "ALBはヘッダーの200 OKを中継しているため、ログ上はelb_status_code=200として記録される。",
            "★★☆☆☆\n(稀なバグ)",
            "レスポンスヘッダーのContent-Lengthと、curl等で直叩きした際の実際のデータサイズに差異がある。",
            "バックエンドアプリケーションのHTTPレスポンス生成ライブラリのバグ修正。手動ヘッダー付与をやめフレームワークに任せる。"
        ),
        (
            "6",
            "【エッジ側のタイムアウト】\nEdge-Optimized CloudFrontの早期タイムアウト・客側切断",
            "REST APIが『Edge-Optimized』の場合、前面にCloudFrontが存在。クライアント側が短時間で通信切断（ブラウザAbort）したか、CloudFrontとAPI GW間の接続エラーでCloudFrontが504を即時返却。",
            "API GatewayとVPC Link v2、ALB、バックエンド間は正常に通信が完了し、ALBは200 OKを記録した（クライアントに届く前に前段で切れた）。",
            "★★★☆☆\n(フロント起因)",
            "504レスポンスのレスポンスヘッダーに『Server: CloudFront』『X-Cache: Error from cloudfront』が含まれる。API GWアクセスログのstatusと比較。",
            "クライアント側のAbortタイムアウト設定を見直す。API Gatewayを『Regional』エンドポイントに変更して構成を明瞭化。"
        ),
        (
            "7",
            "【ネットワーク非対称性】\nVPC Link v2 ENIとALB間の戻りパケット破棄 (NACL)",
            "VPC Link v2のHyperplane ENIからALBへの行き通信は通るが、ALBからAPI GW ENIへの戻り通信がNACLのアウトバウンド/インバウンドルール（エフェメラルポート）で遮断・破棄。",
            "ALBはターゲットから200を受信し、送信ソケットに書き込んで処理成功（200）と認識。しかしパケットがネットワーク上で消滅し、API GW側で即時接続断検知。",
            "★★☆☆☆\n(NW設定時)",
            "VPCフローログでALBからHyperplane ENIへのパケットが『REJECT』されている。特定サブネットのNACLが厳格すぎる。",
            "Hyperplane ENIが配置されたサブネットおよびALBサブネットのNACLで、エフェメラルポート（1024-65535）双方向通信を許可。"
        )
    ]

    for item in all_cases_matrix:
        ws4.row_dimensions[row_idx].height = 68
        for idx_c, val in enumerate(item):
            col_letter = get_column_letter(idx_c + 1)
            cell = ws4[f"{col_letter}{row_idx}"]
            cell.value = val
            cell.font = font_tb
            cell.fill = fill_zebra if (row_idx % 2 == 1) else fill_white
            cell.border = border_thin
            if idx_c in [0, 4]:
                cell.alignment = align_center
            else:
                cell.alignment = align_left
        row_idx += 1

    for idx_c, w in enumerate(col_widths4):
        col_letter = get_column_letter(idx_c + 1)
        ws4.column_dimensions[col_letter].width = w

    # -------------------------------------------------------------
    # SHEET 5: 05_各ケース詳細とログ分析
    # -------------------------------------------------------------
    ws5 = wb.create_sheet(title="05_各ケース詳細とログ分析")
    ws5.views.sheetView[0].showGridLines = True

    # Title Block
    ws5.merge_cells("A1:G1")
    ws5["A1"] = "第5章：各ケースの徹底深堀り・ログ解析と具体的修正手順"
    ws5["A1"].font = font_title
    ws5["A1"].fill = fill_title
    ws5["A1"].alignment = align_center
    ws5.row_dimensions[1].height = 40

    ws5.merge_cells("A2:G2")
    ws5["A2"] = "〜 実際のCloudWatch Execution Log と ALB Access Log で見分ける完全マニュアル 〜"
    ws5["A2"].font = font_subtitle
    ws5["A2"].fill = fill_title
    ws5["A2"].alignment = align_center
    ws5.row_dimensions[2].height = 24

    ws5.row_dimensions[3].height = 15

    row_idx = 4
    
    # Detailed Case 1
    ws5.merge_cells(f"A{row_idx}:G{row_idx}")
    ws5[f"A{row_idx}"] = "【ケース1詳細】API Gateway統合タイムアウト（timeoutInMillis）の極小設定（最頻出！）"
    ws5[f"A{row_idx}"].font = font_section
    ws5[f"A{row_idx}"].fill = fill_section
    ws5[f"A{row_idx}"].alignment = align_left
    ws5.row_dimensions[row_idx].height = 28
    row_idx += 1

    c1_content = [
        ("メカニズム深堀り", "API Gateway REST API のリソース定義において、統合タイムアウトの単位は【ミリ秒 (milliseconds)】です。Terraformの『timeout_in_millis』やCloudFormationで、開発者が『29秒』のつもりで『29』と記述してしまうと、API Gatewayの仕様上の下限値である『50ms』に自動補正されるか、そのまま50msとして保存されます。この結果、API Gatewayはリクエスト送信からわずか0.05秒でタイムアウトを発火させ、クライアントに即時504を返します。バックエンド（ECS）は通常100ms〜200ms程度で処理を終えてALBに200を返し、ALBは『大成功(200)』と記録します。"),
        ("API Gateway ログ例", "(Execution Log)\nHTTP Method: GET, Resource Path: /items\nMethod request body before transformations: ...\nEndpoint request URI: http://my-alb.internal/items\nEndpoint request headers: {x-amzn-apigateway-api-id=..., ...}\nExecution failed due to a timeout: Endpoint request timed out\nMethod completed with status: 504\nLatency: 54 ms (わずか54ミリ秒で終了している！)"),
        ("ALB アクセスログ例", "http 2026-10-07T04:20:00.123456Z app/my-alb/... 10.0.1.50:49152 10.0.2.100:8080 0.001 0.125 0.000 200 200 450 1200 \"GET http://my-alb.internal:80/items HTTP/1.1\"\n→ target_processing_time: 0.125 (125msかかっている)\n→ elb_status_code: 200, target_status_code: 200 (バックエンド側は完全に成功！)"),
        ("修正方法・コード例", "【Terraformの場合の修正】\nresource \"aws_api_gateway_integration\" \"my_integration\" {\n  rest_api_id = aws_api_gateway_rest_api.my_api.id\n  resource_id = aws_api_gateway_resource.my_resource.id\n  http_method = aws_api_gateway_method.my_method.http_method\n  type                    = \"HTTP_PROXY\"\n  integration_http_method = \"GET\"\n  connection_type         = \"VPC_LINK\"\n  connection_id           = aws_api_gateway_vpc_link.v2_link.id\n  uri                     = \"http://${aws_lb.my_alb.dns_name}/items\"\n  timeout_in_millis       = 29000  # ← ★必ずミリ秒(29000)にする！(29と書かない！)\n}")
    ]

    for title, text in c1_content:
        ws5.row_dimensions[row_idx].height = 24
        ws5.merge_cells(f"A{row_idx}:B{row_idx}")
        ws5[f"A{row_idx}"] = title
        ws5[f"A{row_idx}"].font = font_subsection
        ws5[f"A{row_idx}"].fill = fill_highlight
        ws5[f"A{row_idx}"].alignment = align_center
        ws5[f"A{row_idx}"].border = border_thin

        ws5.merge_cells(f"C{row_idx}:G{row_idx}")
        ws5[f"C{row_idx}"] = text
        ws5[f"C{row_idx}"].font = font_code if "Log" in title or "コード" in title else font_tb
        ws5[f"C{row_idx}"].fill = fill_white
        ws5[f"C{row_idx}"].alignment = align_top_left
        ws5[f"C{row_idx}"].border = border_thin
        
        # Adjust height based on text lines
        line_count = text.count("\n") + 1
        ws5.row_dimensions[row_idx].height = max(36, line_count * 18 + 12)
        row_idx += 1

    row_idx += 1
    # Detailed Case 2
    ws5.merge_cells(f"A{row_idx}:G{row_idx}")
    ws5[f"A{row_idx}"] = "【ケース2詳細】調査の錯覚・別人事件（ヘルスチェックや別リクエストの200を見ている）"
    ws5[f"A{row_idx}"].font = font_section
    ws5[f"A{row_idx}"].fill = fill_section
    ws5[f"A{row_idx}"].alignment = align_left
    ws5.row_dimensions[row_idx].height = 28
    row_idx += 1

    c2_content = [
        ("メカニズム深堀り", "障害調査の現場で極めて頻発する『人間の思い込み・誤認』です。クライアントから叩いたリクエストは、VPC Link v2のセキュリティグループ拒否、ALBリスナールール不一致、プライベート名前解決失敗等により、ALBに到達すらしていません。API Gatewayは接続失敗（Connection Refused / RST）を即座に検知し、0〜5ms程度で504（または502）を返しました。一方、エンジニアがALBのCloudWatchログやS3ログを開いた際、ALBターゲットグループが定期実行（30秒毎）している『GET /health』の200ログや、UTC/JSTを混同した別時間のリクエストを見て『ALBは200を返している』と誤認してしまいます。"),
        ("見分け方 (決定打)", "【X-Amzn-Trace-Id の完全突合】\nAPI Gatewayのアクセスログに記録された `$context.extendedRequestId` および `$context.identity.apiKey`、そして何よりも HTTPヘッダーに含まれる `X-Amzn-Trace-Id: Root=1-67035a00-12345678abcdef...` を取得します。\nこのRoot IDでALBアクセスログを `grep` します。もしALBログに該当のTrace-Idが1行もヒットしなければ、【リクエストはALBに一切到達していない】ことが100%確定します！"),
        ("ALBに届いていない場合の真の原因", "① VPC Link v2 のセキュリティグループのアウトバウンドで、ALBのリスナーポート（例: 80 / 443）が許可されていない。\n② ALBのセキュリティグループのインバウンドで、VPC Link v2のセキュリティグループからの通信が許可されていない。\n③ API Gatewayの統合URIで指定したALBのDNS名やプライベートIP、ポート番号が間違っている。\n④ ALBのリスナールール（パスパターンやHostヘッダー条件）に一致せず、デフォルトアクション（固定レスポンス等）にも引っかからない。"),
        ("修正方法", "1. VPC Link v2 に割り当てたセキュリティグループのアウトバウンドルールを確認し、ALBのSGへの通信を許可。\n2. ALBのセキュリティグループのインバウンドルールに、VPC Link v2のセキュリティグループIDをソースとしてポート80/443を許可。\n3. API Gateway統合リクエストのURI（例: `http://${alb_dns}:80/api/...`）が正確であることを確認。")
    ]

    for title, text in c2_content:
        ws5.row_dimensions[row_idx].height = 24
        ws5.merge_cells(f"A{row_idx}:B{row_idx}")
        ws5[f"A{row_idx}"] = title
        ws5[f"A{row_idx}"].font = font_subsection
        ws5[f"A{row_idx}"].fill = fill_highlight
        ws5[f"A{row_idx}"].alignment = align_center
        ws5[f"A{row_idx}"].border = border_thin

        ws5.merge_cells(f"C{row_idx}:G{row_idx}")
        ws5[f"C{row_idx}"] = text
        ws5[f"C{row_idx}"].font = font_tb
        ws5[f"C{row_idx}"].fill = fill_white
        ws5[f"C{row_idx}"].alignment = align_top_left
        ws5[f"C{row_idx}"].border = border_thin
        
        line_count = text.count("\n") + 1
        ws5.row_dimensions[row_idx].height = max(36, line_count * 18 + 12)
        row_idx += 1

    row_idx += 1
    # Detailed Case 3 & 4
    ws5.merge_cells(f"A{row_idx}:G{row_idx}")
    ws5[f"A{row_idx}"] = "【ケース3 & 4詳細】Keep-Alive競合レースコンディション / チャンク転送の異常切断"
    ws5[f"A{row_idx}"].font = font_section
    ws5[f"A{row_idx}"].fill = fill_section
    ws5[f"A{row_idx}"].alignment = align_left
    ws5.row_dimensions[row_idx].height = 28
    row_idx += 1

    c3_c4_content = [
        ("Keep-Alive競合の深堀り", "API Gatewayは接続パフォーマンス向上のため、VPC Link経由でALBへのTCPセッションを一定時間維持（Keep-Alive）して再利用します。ALBのデフォルトのアイドルタイムアウトは60秒です。直前の通信からちょうど60秒が経過する瞬間、ALBはTCP FINパケットを送出してソケットをクローズしようとします。全く同じ瞬間にAPI Gatewayが新しいリクエストパケットをそのソケットへ送信すると、ALBは切断途中のソケットにデータが届いたため『TCP RST』を返送します。API Gatewayはリクエスト送信直後に接続が強制リセットされたため、即時エラーをクライアントに返します。"),
        ("Keep-Aliveの対策", "ALB側のアイドルタイムアウト（Idle Timeout）設定値を、API Gateway側のキープアライブ維持時間（非公開だが通常数十秒〜60秒未満）よりも【意図的に長く（例: 65秒〜120秒）】設定します。これにより、クライアント側（API Gateway）が先に安全に接続を閉じるように主導権を握らせ、ALB側からの不意なTCP FIN送信による衝突を防ぎます。"),
        ("チャンク転送異常の深堀り", "バックエンドが『Transfer-Encoding: chunked』でレスポンスを返す場合、ALBはヘッダーを受信した時点で200ステータスを認識し、ALBログにもelb_status_code=200を出力します。しかし、バックエンドがボディを送信している最中にOOM（メモリ不足）等で異常終了すると、終端チャンクが送られないままTCP接続が切断されます。REST API Gatewayは完全バッファリング型プロキシであるため、レスポンスが途中で壊れたと判断し、クライアントにエラー（504または502）を即座に返します。"),
        ("チャンク転送の対策", "① バックエンドアプリケーションのメモリ割り当てを増やし、コンテナのOOM Killer死を防止する。\n② REST APIの前段に置くアプリケーションでは、可能であれば完全な『Content-Length』を付与して一括レスポンスを返す設計にする。\n③ 大容量ストリーミングが必要な場合は、REST APIではなくストリーミングに対応した『HTTP API』への移行を検討する。")
    ]

    for title, text in c3_c4_content:
        ws5.row_dimensions[row_idx].height = 24
        ws5.merge_cells(f"A{row_idx}:B{row_idx}")
        ws5[f"A{row_idx}"] = title
        ws5[f"A{row_idx}"].font = font_subsection
        ws5[f"A{row_idx}"].fill = fill_highlight
        ws5[f"A{row_idx}"].alignment = align_center
        ws5[f"A{row_idx}"].border = border_thin

        ws5.merge_cells(f"C{row_idx}:G{row_idx}")
        ws5[f"C{row_idx}"] = text
        ws5[f"C{row_idx}"].font = font_tb
        ws5[f"C{row_idx}"].fill = fill_white
        ws5[f"C{row_idx}"].alignment = align_top_left
        ws5[f"C{row_idx}"].border = border_thin
        
        line_count = text.count("\n") + 1
        ws5.row_dimensions[row_idx].height = max(36, line_count * 18 + 12)
        row_idx += 1

    ws5.column_dimensions["A"].width = 16
    ws5.column_dimensions["B"].width = 16
    ws5.column_dimensions["C"].width = 24
    ws5.column_dimensions["D"].width = 24
    ws5.column_dimensions["E"].width = 24
    ws5.column_dimensions["F"].width = 24
    ws5.column_dimensions["G"].width = 24

    # -------------------------------------------------------------
    # SHEET 6: 06_トラブルシューティング手順
    # -------------------------------------------------------------
    ws6 = wb.create_sheet(title="06_トラブルシューティング手順")
    ws6.views.sheetView[0].showGridLines = True

    # Title Block
    ws6.merge_cells("A1:G1")
    ws6["A1"] = "第6章：真の原因を5分で特定する調査フローチャートとチェックリスト"
    ws6["A1"].font = font_title
    ws6["A1"].fill = fill_title
    ws6["A1"].alignment = align_center
    ws6.row_dimensions[1].height = 40

    ws6.merge_cells("A2:G2")
    ws6["A2"] = "〜 現場のエンジニアが迷わず即座に根本原因へ到達するための実戦的診断ガイド 〜"
    ws6["A2"].font = font_subtitle
    ws6["A2"].fill = fill_title
    ws6["A2"].alignment = align_center
    ws6.row_dimensions[2].height = 24

    ws6.row_dimensions[3].height = 15

    row_idx = 4
    ws6.merge_cells(f"A{row_idx}:G{row_idx}")
    ws6[f"A{row_idx}"] = "■ 5分で原因を特定するトラブルシューティング・ディシジョンツリー"
    ws6[f"A{row_idx}"].font = font_section
    ws6[f"A{row_idx}"].fill = fill_section
    ws6[f"A{row_idx}"].alignment = align_left
    ws6.row_dimensions[row_idx].height = 28
    row_idx += 1

    tree_headers = ["ステップ", "確認アクション", "判定基準 (IF)", "診断結果 (THEN)", "次に打つべき手 (ACTION)"]
    ws6.row_dimensions[row_idx].height = 26
    ws6.merge_cells(f"D{row_idx}:E{row_idx}")
    ws6.merge_cells(f"F{row_idx}:G{row_idx}")
    
    ws6[f"A{row_idx}"] = "ステップ"
    ws6[f"B{row_idx}"] = "確認アクション"
    ws6[f"C{row_idx}"] = "判定基準 (IF)"
    ws6[f"D{row_idx}"] = "診断結果 (THEN)"
    ws6[f"F{row_idx}"] = "次に打つべき手 (ACTION)"

    for col_c in ["A", "B", "C", "D", "E", "F", "G"]:
        ws6[f"{col_c}{row_idx}"].font = font_th
        ws6[f"{col_c}{row_idx}"].fill = fill_th
        ws6[f"{col_c}{row_idx}"].alignment = align_center
        ws6[f"{col_c}{row_idx}"].border = border_th
    row_idx += 1

    tree_data = [
        ("Step 1\nTrace ID突合", "API GW実行ログの x-amzn-trace-id をALBログで完全一致検索", "ALBログに一致するTrace IDが【存在しない】", "【ケース2】リクエスト取り違え\n実際のリクエストはALBに届いていない", "VPC Link v2とALBのSecurity Group、リスナーURL設定を点検せよ"),
        ("Step 2\nTrace ID一致時", "ALBログにTrace IDが存在する場合、API GWのレイテンシを確認", "API Gatewayログの Latency が【50ms〜500ms】で極小", "【ケース1】統合タイムアウト誤設定\ntimeoutInMillis が極小値になっている", "Terraformの timeout_in_millis を 29000 (29秒) に修正して再デプロイ"),
        ("Step 3\nタイムアウト正常時", "ALBアクセスログの target_processing_time と sent_bytes を確認", "elb_status_code=200 だが sent_bytes が極端に小さい / 0", "【ケース4 / 5】ボディ途絶・プロトコル異常\nChunked送信途中でプロセス死亡", "バックエンドコンテナのOOMログ確認、例外処理の修正を実施"),
        ("Step 4\nエラー断続発生時", "エラーの発生タイミングと直前通信の間隔をログ分析", "直前の通信から約60秒後などアイドル直後にのみ発生", "【ケース3】Keep-Aliveレースコンディション\nALBアイドルTOとAPI GWプールの衝突", "ALBのアイドルタイムアウトを 65秒〜120秒 に延ばしてAPI GWより長寿命化"),
        ("Step 5\nレスポンスヘッダ確認", "クライアントが受け取ったレスポンスヘッダーの Server を確認", "Server: CloudFront / X-Cache: Error from cloudfront", "【ケース6】Edge側タイムアウト\n手前のCloudFrontが切断している", "クライアント側のタイムアウト見直し、Regionalエンドポイントへの移行検討")
    ]

    for item in tree_data:
        ws6.row_dimensions[row_idx].height = 48
        ws6.merge_cells(f"D{row_idx}:E{row_idx}")
        ws6.merge_cells(f"F{row_idx}:G{row_idx}")
        
        ws6[f"A{row_idx}"] = item[0]
        ws6[f"B{row_idx}"] = item[1]
        ws6[f"C{row_idx}"] = item[2]
        ws6[f"D{row_idx}"] = item[3]
        ws6[f"F{row_idx}"] = item[4]
        
        fill_curr = fill_zebra if (row_idx % 2 == 1) else fill_white
        for col_c in ["A", "B", "C", "D", "E", "F", "G"]:
            cell = ws6[f"{col_c}{row_idx}"]
            cell.font = font_tb
            cell.fill = fill_curr
            cell.border = border_thin
            if col_c in ["A", "C"]:
                cell.alignment = align_center
            else:
                cell.alignment = align_left
        row_idx += 1

    row_idx += 1
    # Final Checklist
    ws6.merge_cells(f"A{row_idx}:G{row_idx}")
    ws6[f"A{row_idx}"] = "■ 本番環境で今すぐ確認すべき設定チェックリスト"
    ws6[f"A{row_idx}"].font = font_section
    ws6[f"A{row_idx}"].fill = fill_section
    ws6[f"A{row_idx}"].alignment = align_left
    ws6.row_dimensions[row_idx].height = 28
    row_idx += 1

    check_headers = ["対象コンポーネント", "設定項目 / パラメータ", "推奨設定値 / 正常基準", "NG例 (即時504を招く設定)"]
    ws6.row_dimensions[row_idx].height = 26
    ws6.merge_cells(f"D{row_idx}:G{row_idx}")
    
    ws6[f"A{row_idx}"] = "対象コンポーネント"
    ws6[f"B{row_idx}"] = "設定項目 / パラメータ"
    ws6[f"C{row_idx}"] = "推奨設定値 / 正常基準"
    ws6[f"D{row_idx}"] = "NG例 (即時504を招く設定)"

    for col_c in ["A", "B", "C", "D", "E", "F", "G"]:
        ws6[f"{col_c}{row_idx}"].font = font_th
        ws6[f"{col_c}{row_idx}"].fill = fill_th
        ws6[f"{col_c}{row_idx}"].alignment = align_center
        ws6[f"{col_c}{row_idx}"].border = border_th
    row_idx += 1

    checks = [
        ("API Gateway Method Integration", "timeoutInMillis", "29000 (デフォルト) または 5000〜10000ms", "29, 50, 100 (ミリ秒と秒を勘違いした極小値)"),
        ("ALB (Load Balancer Attribute)", "idle_timeout.timeout_seconds", "65秒 〜 120秒 (API GWより長い値)", "30秒以下、またはAPI GWキープアライブ時間と完全同値"),
        ("ALB Security Group", "Inbound Rules (Port 80/443)", "VPC Link v2 の Security Group ID を許可", "特定のプライベートIPのみ許可、またはSG許可漏れ"),
        ("VPC Link v2 Security Group", "Outbound Rules", "ALBのセキュリティグループまたはポートを許可", "アウトバウンドが全拒否、または別ポートのみ許可"),
        ("Subnet Network ACL (NACL)", "Inbound / Outbound", "エフェメラルポート (1024-65535) 双方向許可", "戻りパケットポート（1024-65535）がDENYされている"),
        ("Backend Application", "Response Formatting", "完全なContent-Lengthヘッダーの付与", "Transfer-Encoding: chunked 送信途中のプロセス強制終了")
    ]

    for item in checks:
        ws6.row_dimensions[row_idx].height = 36
        ws6.merge_cells(f"D{row_idx}:G{row_idx}")
        ws6[f"A{row_idx}"] = item[0]
        ws6[f"B{row_idx}"] = item[1]
        ws6[f"C{row_idx}"] = item[2]
        ws6[f"D{row_idx}"] = item[3]
        
        fill_curr = fill_zebra if (row_idx % 2 == 1) else fill_white
        for col_c in ["A", "B", "C", "D", "E", "F", "G"]:
            cell = ws6[f"{col_c}{row_idx}"]
            cell.font = font_tb
            cell.fill = fill_curr
            cell.border = border_thin
            if col_c in ["A", "B"]:
                cell.alignment = align_center
            else:
                cell.alignment = align_left
        row_idx += 1

    ws6.column_dimensions["A"].width = 24
    ws6.column_dimensions["B"].width = 28
    ws6.column_dimensions["C"].width = 32
    ws6.column_dimensions["D"].width = 26
    ws6.column_dimensions["E"].width = 26
    ws6.column_dimensions["F"].width = 26
    ws6.column_dimensions["G"].width = 26

    # Save workbook
    output_filename = "API_Gateway_504_Immediate_RootCause_Guide.xlsx"
    wb.save(output_filename)
    print(f"Successfully generated: {output_filename}")

if __name__ == "__main__":
    create_workbook()
