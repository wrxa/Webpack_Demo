module.exports = {
    devtool: 'eval-source-map',
    entry: __dirname + "/app/main.js", //メインJSファイルパス
    output: {
        path: __dirname + "/public", //パッケージしたJSファイルの格納フォルダ
        filename: "bundle.js" //パッケージしたJSファイル
    },
    devServer: { // 内蔵サーバ設定
        contentBase: "./public", //サーバファイルフォルダ（index.html、資源ファイル、bundleJSなど）
        historyApiFallback: true, //
        inline: true, //ソースコードの修正に伴い、WEBページをリアルタイム更新する
        // host: '0.0.0.0' //IPアドレス形式でもサーバをアクセスできようにする（例えば：10.167.20.3）
    },
    // babel設定
    module: {
        rules: [{
	            test: require.resolve('jquery'), //require.resolve 用来获取模块的绝对路径
	            use: [{
	                loader: 'expose-loader',
	                options: 'jQuery'
	            }, {
	                loader: 'expose-loader',
	                options: '$'
	            }]
	        }, {
                test: /(\.jsx|\.js)$/,
                exclude: /(node_modules)/,
                use: [{
                    loader: 'babel-loader',
                    query: {
                        presets: ["react"]
                    }
                }]
            }, {
                test: /\.css$/,
                use: ['style-loader', 'css-loader']
            }
        ]
    }
}