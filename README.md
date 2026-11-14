# skills
Agent Skills

## スキル一覧

| スキル | 説明 |
| --- | --- |
| [windows-speak](windows-speak/) | Windows標準の音声合成(System.Speech)で、文章の読み上げ、タスク完了の通知、ユーザーへの説明を声で行う。Claudeが読み上げ用の文章を別に作って話す。導入時のCLAUDE.mdへの追記とStop hookの設定は [windows-speak/README.md](windows-speak/README.md) を参照。 |
| [repository-html-browser](repository-html-browser/) | リポジトリ全体を探索し、各ファイルをVS Code風の閲覧用HTMLへ変換して `html-code/` に出力する。元のディレクトリ構造を保持し、`フォルダーツリー.html` から全ファイルへ移動できる。変換は同梱の `scripts/generate-html-code.mjs` が行う。 |
