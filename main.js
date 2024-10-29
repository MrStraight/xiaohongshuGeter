const { app, BrowserWindow, ipcMain, dialog, shell } = require('electron');
const path = require('path');
const fs = require('fs').promises;
const { scrapeXiaohongshu } = require('./demo.js');
const puppeteer = require('puppeteer');

function createWindow() {
    const win = new BrowserWindow({
        width: 1200,  // 增加窗口宽度
        height: 800,  // 增加窗口高度
        webPreferences: {
            nodeIntegration: true,
            contextIsolation: false
        },
        // 添加以下配置来移除菜单栏
        autoHideMenuBar: true,  // 自动隐藏菜单栏
        menuBarVisible: false,   // 不显示菜单栏
    });

    // 移除默认菜单
    win.removeMenu();

    win.loadFile('index.html');
}

app.whenReady().then(createWindow);

app.on('window-all-closed', () => {
    if (process.platform !== 'darwin') {
        app.quit();
    }
});

app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) {
        createWindow();
    }
});

// 添加停止标志
let shouldStopScraping = false;

// 修改抓取请求处理
ipcMain.on('start-scrape', async (event, data) => {
    try {
        shouldStopScraping = false;
        event.reply('update-status', '开始抓取...');
        
        // 创建一个函数来处理新文章
        const handleNewArticle = (article) => {
            if (shouldStopScraping) return false;  // 如果需要停止，返回 false
            
            event.reply('new-article', {
                data: {
                    article: article,
                    totalCount: 1
                }
            });
            return true;  // 继续抓取
        };

        // 创建一个函数来处理进度更新
        const handleProgress = (progress) => {
            if (shouldStopScraping) return false;  // 如果需要停止，返回 false
            event.reply('update-progress', progress);
            return true;  // 继续抓取
        };

        // 调用抓取函数，传入所有参数
        const articles = await scrapeXiaohongshu(
            data.keyword, 
            handleNewArticle, 
            handleProgress,
            data.sortType,
            data.cookieIndex  // 传入选中的 cookie 索引
        );
        
        event.reply('scrape-complete', articles);
    } catch (error) {
        console.error('抓取错误：', error);
        event.reply('scrape-error', error.message);
    }
});

// 添加停止抓取的处理
ipcMain.on('stop-scrape', (event) => {
    shouldStopScraping = true;
});

// 修改导出处理部分
ipcMain.on('export-data', async (event, articles) => {
    try {
        // 获取当前搜索关键词
        const keyword = articles[0]?.keyword || '小红书文章';  // 如果没有关键词就使用默认名称
        const currentTime = new Date().toLocaleString('zh-CN', {
            year: 'numeric',
            month: '2-digit',
            day: '2-digit',
            hour: '2-digit',
            minute: '2-digit',
            second: '2-digit'
        }).replace(/[\/\s:]/g, '');  // 移除日期时间中的特殊字符

        // 生成文件名
        const defaultFileName = `${keyword}-${currentTime}`;

        const { filePath, canceled } = await dialog.showSaveDialog({
            filters: [
                { name: 'CSV 文件', extensions: ['csv'] },
                { name: 'Excel 文件', extensions: ['xlsx'] },
                { name: 'JSON 文件', extensions: ['json'] }
            ],
            defaultPath: path.join(app.getPath('downloads'), defaultFileName)
        });

        if (canceled) return;

        const ext = path.extname(filePath).toLowerCase();
        
        switch (ext) {
            case '.csv':
                await exportToCsv(articles, filePath);
                break;
            case '.xlsx':
                await exportToExcel(articles, filePath);
                break;
            case '.json':
                await exportToJson(articles, filePath);
                break;
        }

        event.reply('export-complete', filePath);
    } catch (error) {
        event.reply('export-error', error.message);
    }
});

async function exportToCsv(articles, filePath) {
    const headers = ['序号', '标题', '作者', '作者链接', '点赞数', '链', '时间'];
    const rows = articles.map((article, index) => [
        index + 1,
        article.title,
        article.author,
        article.authorUrl || '',  // 添加作者链接
        article.likes,
        article.url,
        new Date(article.timestamp).toLocaleString()
    ]);
    
    const csvContent = [
        headers.join(','),
        ...rows.map(row => row.map(cell => `"${(cell || '').toString().replace(/"/g, '""')}"`).join(','))
    ].join('\n');
    
    await fs.writeFile(filePath, '\ufeff' + csvContent, 'utf8');
}

async function exportToJson(articles, filePath) {
    await fs.writeFile(filePath, JSON.stringify(articles, null, 2), 'utf8');
}

async function exportToExcel(articles, filePath) {
    const XLSX = require('xlsx');
    
    const ws = XLSX.utils.json_to_sheet(articles.map((article, index) => ({
        '序号': index + 1,
        '标题': article.title,
        '作者': article.author,
        '作者链接': article.authorUrl || '',  // 添加作者链接
        '点赞数': article.likes,
        '链接': article.url,
        '时间': new Date(article.timestamp).toLocaleString()
    })));
    
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, '相亲文章');
    
    XLSX.writeFile(wb, filePath);
}

// 修改 cookie 保存理
ipcMain.on('save-cookie', async (event, data) => {
    try {
        const { cookieString, note } = data;
        const cookies = cookieString.split(';').map(cookie => {
            const [name, value] = cookie.trim().split('=');
            return {
                name,
                value,
                domain: '.xiaohongshu.com'
            };
        });

        // 使用 app.getPath('userData') 来获取应用数据目录
        const configPath = path.join(app.getPath('userData'), 'config.js');
        
        // 读取现有配置
        let config = { cookiesList: [], cookieNotes: [] };
        try {
            // 检查文件是否存在
            try {
                await fs.access(configPath);
                const configContent = await fs.readFile(configPath, 'utf8');
                if (configContent) {
                    // 使用 JSON 格式而不是 JavaScript 模块格式
                    config = JSON.parse(configContent);
                }
            } catch (error) {
                console.log('配置文件不存在或解析失败，使用默认配置');
            }
        } catch (error) {
            console.log('创建新的配置文件');
        }

        // 确保数组存在
        if (!Array.isArray(config.cookiesList)) config.cookiesList = [];
        if (!Array.isArray(config.cookieNotes)) config.cookieNotes = [];

        // 添加新的 cookie 组和备注到列表中
        config.cookiesList.push(cookies);
        config.cookieNotes.push(note || '');

        // 直接保存为 JSON 格式
        await fs.writeFile(configPath, JSON.stringify(config, null, 2), 'utf8');

        // 返回更新后的数据
        event.reply('cookie-saved', { 
            success: true, 
            cookiesList: config.cookiesList,
            cookieNotes: config.cookieNotes,
            cookiesCount: config.cookiesList.length 
        });

        console.log('Cookie 保存成功，当前共有', config.cookiesList.length, '组');
    } catch (error) {
        console.error('保存 cookie 失败:', error);
        event.reply('cookie-saved', { success: false });
    }
});

// 修改清除 cookie 处理
ipcMain.on('clear-cookie', async (event) => {
    try {
        // 使用 app.getPath('userData') 来获取应用数据目录
        const configPath = path.join(app.getPath('userData'), 'config.js');
        
        // 创建一个空的配置文件
        const configContent = `module.exports = {
    cookiesList: [],
    cookieNotes: []
};`;

        // 写入空配置
        await fs.writeFile(configPath, configContent, 'utf8');
        
        // 成功
        event.reply('cookie-cleared', { 
            success: true,
            cookiesList: [],
            cookieNotes: [],
            cookiesCount: 0
        });
    } catch (error) {
        console.error('清除 cookie 失败:', error);
        event.reply('cookie-cleared', { success: false });
    }
});

// 修改打开文章的处理
ipcMain.on('open-article', async (event, url) => {
    try {
        // 确保使用正确的链接格式
        let finalUrl = url;
        const noteId = url.match(/\/(?:explore|discovery\/item)\/(\w+)/)?.[1];
        if (noteId) {
            finalUrl = `https://www.xiaohongshu.com/discovery/item/${noteId}`;
        }

        // 使用默认浏览器打开链接
        await shell.openExternal(finalUrl);
    } catch (error) {
        console.error('打开文章失败:', error);
    }
});

// 修改获取 cookie 列表的处理
ipcMain.on('get-cookies', async (event) => {
    try {
        const configPath = path.join(app.getPath('userData'), 'config.js');
        const configContent = await fs.readFile(configPath, 'utf8');
        // 使用 JSON.parse 而不是 eval
        const config = JSON.parse(configContent);
        
        event.reply('cookies-updated', {
            cookiesList: config.cookiesList || [],
            cookieNotes: config.cookieNotes || [],
            cookiesCount: config.cookiesList?.length || 0
        });
    } catch (error) {
        console.error('获取 cookie 失败:', error);
        // 如果文件不存在或解析失败，尝试从默认配置文件读取xx
        try {
            const defaultConfig = require('./config.js');
            event.reply('cookies-updated', {
                cookiesList: defaultConfig.cookiesList || [],
                cookieNotes: defaultConfig.cookieNotes || [],
                cookiesCount: defaultConfig.cookiesList?.length || 0
            });
        } catch (defaultError) {
            console.error('读取默认配置失败:', defaultError);
            event.reply('cookies-updated', { 
                cookiesList: [], 
                cookieNotes: [],
                cookiesCount: 0 
            });
        }
    }
});

// 修改 cookie 检查处理
ipcMain.on('check-cookies', async (event) => {
    try {
        const configPath = path.join(app.getPath('userData'), 'config.js');
        const configContent = await fs.readFile(configPath, 'utf8');
        // 使用 JSON.parse 而不是 eval
        const config = JSON.parse(configContent);
        
        // 检查 cookiesList 是否存在且有内容
        const hasCookies = config && config.cookiesList && Array.isArray(config.cookiesList) && config.cookiesList.length > 0;
        console.log('检查 Cookie 结果:', hasCookies, '当前 Cookie 数量:', config?.cookiesList?.length || 0);
        event.reply('cookies-check-result', hasCookies);
    } catch (error) {
        console.error('检查 cookie 失败:', error);
        // 如果文件不存在或解析失败，尝试从默认配置文件读取
        try {
            const defaultConfig = require('./config.js');
            const hasCookies = defaultConfig && defaultConfig.cookiesList && Array.isArray(defaultConfig.cookiesList) && defaultConfig.cookiesList.length > 0;
            console.log('使用默认配置检查 Cookie 结果:', hasCookies, '当前 Cookie 数量:', defaultConfig?.cookiesList?.length || 0);
            event.reply('cookies-check-result', hasCookies);
        } catch (defaultError) {
            console.error('读取默认配置失败:', defaultError);
            event.reply('cookies-check-result', false);
        }
    }
});

// 修改删除单个 cookie 组的处理
ipcMain.on('delete-cookie', async (event, index) => {
    try {
        const configPath = path.join(app.getPath('userData'), 'config.js');
        
        // 读取现有配置
        const configContent = await fs.readFile(configPath, 'utf8');
        let config = JSON.parse(configContent);  // 使用 JSON.parse 而不是 eval
        
        if (Array.isArray(config.cookiesList)) {
            // 删除指定索引的 cookie 组和备注
            config.cookiesList.splice(index, 1);
            config.cookieNotes.splice(index, 1);
            
            // 直接保存为 JSON 格式
            await fs.writeFile(configPath, JSON.stringify(config, null, 2), 'utf8');
            
            event.reply('cookie-deleted', { 
                success: true, 
                cookiesList: config.cookiesList,
                cookieNotes: config.cookieNotes,
                cookiesCount: config.cookiesList.length 
            });
        } else {
            throw new Error('Cookie 列表格式错误');
        }
    } catch (error) {
        console.error('删除 cookie 失败:', error);
        event.reply('cookie-deleted', { success: false });
    }
});

// 添加打开作者主页的处理
ipcMain.on('open-author-page', async (event, url) => {
    try {
        // 确保使用正确链接格式
        let finalUrl = url;
        if (url.startsWith('/user/profile/')) {
            finalUrl = `https://www.xiaohongshu.com${url}`;
        }

        // 使用默浏览器打开链接
        await shell.openExternal(finalUrl);
    } catch (error) {
        console.error('打开作者主页失败:', error);
    }
});

// 添加导入数据处理
ipcMain.on('import-data', async (event) => {
    try {
        const { filePaths, canceled } = await dialog.showOpenDialog({
            filters: [
                { name: 'CSV 文件', extensions: ['csv'] },
                { name: 'Excel 文件', extensions: ['xlsx'] },
                { name: 'JSON 文件', extensions: ['json'] }
            ],
            properties: ['openFile']
        });

        if (canceled || filePaths.length === 0) return;

        const filePath = filePaths[0];
        const ext = path.extname(filePath).toLowerCase();
        let articles;

        switch (ext) {
            case '.csv':
                articles = await importFromCsv(filePath);
                break;
            case '.xlsx':
                articles = await importFromExcel(filePath);
                break;
            case '.json':
                articles = await importFromJson(filePath);
                break;
            default:
                throw new Error('不支持的文件格式');
        }

        console.log('导入数据:', articles.length, '篇文章');
        event.reply('import-complete', articles);
    } catch (error) {
        console.error('导入错误:', error);
        event.reply('import-error', error.message);
    }
});

// 修改导入函数
async function importFromCsv(filePath) {
    try {
        const content = await fs.readFile(filePath, 'utf8');
        // 移除 BOM 记
        const cleanContent = content.replace(/^\uFEFF/, '');
        
        // 按行分割
        const lines = cleanContent.split('\n').filter(line => line.trim());
        
        // 移除标题
        const headers = lines.shift();
        
        // 解析每一行数据
        const articles = lines.map(line => {
            // 处理引号内的逗号
            const values = line.match(/(".*?"|[^",\s]+)(?=\s*,|\s*$)/g) || [];
            // 清理引号
            const cleanValues = values.map(val => val.replace(/^"|"$/g, '').replace(/""/g, '"'));
            
            // 处理作者链接
            const authorUrl = cleanValues[3] || ''; // 作者链接在第4列
            const formattedAuthorUrl = authorUrl.startsWith('/user/profile/') 
                ? `https://www.xiaohongshu.com${authorUrl}`
                : authorUrl;

            // 处理时间格式
            let timestamp;
            try {
                // 尝试解析时间字符串
                const timeStr = cleanValues[6] || ''; // 时间在第7列
                timestamp = new Date(timeStr).toISOString();
            } catch (error) {
                // 果解失败，使用当前时间
                timestamp = new Date().toISOString();
            }

            return {
                title: cleanValues[1] || '',
                author: cleanValues[2] || '',
                authorUrl: formattedAuthorUrl,
                likes: parseInt(cleanValues[4]) || 0,
                url: cleanValues[5] || '',
                timestamp: timestamp
            };
        }).filter(article => article.title || article.url);

        console.log('成功导入文章数:', articles.length);
        return articles;
    } catch (error) {
        console.error('CSV 导入错误:', error);
        throw new Error('CSV 文件解析失败: ' + error.message);
    }
}

async function importFromJson(filePath) {
    try {
        const content = await fs.readFile(filePath, 'utf8');
        const data = JSON.parse(content);
        
        // 确保数据是数组
        const articles = Array.isArray(data) ? data : [data];
        
        // 规范化数据格式，处理作者链接
        return articles.map(item => {
            const authorUrl = item.authorUrl || '';
            const formattedAuthorUrl = authorUrl.startsWith('/user/profile/') 
                ? `https://www.xiaohongshu.com${authorUrl}`
                : authorUrl;

            return {
                title: item.title || '',
                author: item.author || '',
                likes: parseInt(item.likes) || 0,
                url: item.url || '',
                timestamp: item.timestamp || new Date().toISOString(),
                authorUrl: formattedAuthorUrl
            };
        }).filter(article => article.title || article.url);
    } catch (error) {
        console.error('JSON 导入错误:', error);
        throw new Error('JSON 文件解析失败: ' + error.message);
    }
}

async function importFromExcel(filePath) {
    try {
        const XLSX = require('xlsx');
        const workbook = XLSX.readFile(filePath);
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        const data = XLSX.utils.sheet_to_json(worksheet);

        return data.map(row => {
            const authorUrl = row['作者链接'] || '';
            const formattedAuthorUrl = authorUrl.startsWith('/user/profile/') 
                ? `https://www.xiaohongshu.com${authorUrl}`
                : authorUrl;

            return {
                title: row['标题'] || '',
                author: row['作者'] || '',
                likes: parseInt(row['点赞数']) || 0,
                url: row['链接'] || '',
                timestamp: new Date(row['时间'] || Date.now()).toISOString(),
                authorUrl: formattedAuthorUrl
            };
        }).filter(article => article.title || article.url);
    } catch (error) {
        console.error('Excel 导入错误:', error);
        throw new Error('Excel 文件解析失败: ' + error.message);
    }
}

// 添加停止标志
let shouldStopActions = false;

// 添加停止操作的处理
ipcMain.on('stop-actions', (event) => {
    shouldStopActions = true;
});

// 修改执行操作的处理，添加验证码检查
ipcMain.on('run-actions', async (event, data) => {
    try {
        shouldStopActions = false;
        const browser = await puppeteer.launch({
            headless: false,
            defaultViewport: null,
            args: [
                '--window-size=1920,1080',
                '--disable-web-security',
                '--disable-features=IsolateOrigins,site-per-process'
            ]
        });

        const sleep = (ms) => new Promise(resolve => setTimeout(resolve, ms));

        let successCount = 0;
        let currentCount = 0;
        const totalUrls = data.urls.length;

        // 获取当前选中的 cookie
        const configPath = path.join(app.getPath('userData'), 'config.js');
        const configContent = await fs.readFile(configPath, 'utf8');
        const config = JSON.parse(configContent);

        // 获取当前选中的 cookie
        const selectedCookies = config.cookiesList[data.cookieIndex];
        const cookieNote = config.cookieNotes[data.cookieIndex] || `Cookie组 ${data.cookieIndex + 1}`;
        console.log('\x1b[36m%s\x1b[0m', `使用 Cookie 组: ${data.cookieIndex}, 共 ${selectedCookies?.length || 0} 个 cookies`);

        for (const url of data.urls) {
            if (shouldStopActions) {
                console.log('\x1b[33m%s\x1b[0m', '收到停止信号，终止操作');
                break;
            }

            currentCount++;
            const page = await browser.newPage();
            try {
                // 设置请求头
                await page.setExtraHTTPHeaders({
                    'Accept-Language': 'zh-CN,zh;q=0.9',
                    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8',
                    'Accept-Encoding': 'gzip, deflate, br'
                });

                await page.setUserAgent('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36');

                // 确保在访问页面前设置 cookie
                if (selectedCookies && selectedCookies.length > 0) {
                    for (const cookie of selectedCookies) {
                        await page.setCookie({
                            ...cookie,
                            domain: '.xiaohongshu.com'
                        });
                    }
                    console.log('\x1b[32m%s\x1b[0m', `已设置 Cookie: ${cookieNote}`);
                }

                console.log('\x1b[34m%s\x1b[0m', `正在处理第 ${currentCount}/${totalUrls} 个链接: ${url}`);

                await page.goto(url, { 
                    waitUntil: 'networkidle0',
                    timeout: 30000 
                });

                // 检查验证码
                const hasCaptcha = await page.evaluate(() => {
                    const captchaElement = document.querySelector('.red-captcha-title');
                    return captchaElement && captchaElement.textContent.includes('请通过验证');
                });

                if (hasCaptcha) {
                    console.log('\x1b[31m%s\x1b[0m', `检测到验证码，当前运行到第 ${currentCount}/${totalUrls} 个，还剩 ${totalUrls - currentCount} 个未运行`);
                    event.reply('captcha-detected', {
                        current: currentCount,
                        total: totalUrls,
                        remaining: totalUrls - currentCount
                    });
                    await browser.close();
                    return;
                }

                // 等待页面加载，使用 sleep 替代 waitForTimeout
                await sleep(2000);

                // 更新进度
                event.reply('action-progress', {
                    current: currentCount,
                    total: totalUrls,
                    url: url
                });

                // 执行选中的操作
                if (data.actions.follow) {
                    console.log('\x1b[33m%s\x1b[0m', '尝试点击关注按钮');
                    await page.waitForSelector('button.follow-button', { timeout: 5000 });
                    await page.evaluate(() => {
                        const followBtn = document.querySelector('button.follow-button');
                        if (followBtn) {
                            console.log('找到关注按钮:', followBtn.outerHTML);
                            followBtn.click();
                        }
                    });
                    await sleep(1000);
                }

                if (data.actions.like) {
                    console.log('\x1b[33m%s\x1b[0m', '尝试点击点赞按钮');
                    await page.waitForSelector('.interact-container .like-active', { timeout: 5000 });
                    await page.evaluate(() => {
                        const likeBtn = document.querySelector('.interact-container .like-active');
                        if (likeBtn) {
                            console.log('找到点赞按钮:', likeBtn.outerHTML);
                            likeBtn.click();
                        }
                    });
                    await sleep(1000);
                }

                if (data.actions.favorite) {
                    console.log('\x1b[33m%s\x1b[0m', '尝试点击收藏按钮');
                    await page.waitForSelector('.interact-container #note-page-collect-board-guide', { timeout: 5000 });
                    await page.evaluate(() => {
                        const collectBtn = document.querySelector('.interact-container #note-page-collect-board-guide');
                        if (collectBtn) {
                            console.log('找到收藏按钮:', collectBtn.outerHTML);
                            collectBtn.click();
                        }
                    });
                    await sleep(1000);
                }

                successCount++;
                console.log('\x1b[32m%s\x1b[0m', `成功处理第 ${currentCount}/${totalUrls} 个链接`);
                await page.close();
            } catch (error) {
                console.error('\x1b[31m%s\x1b[0m', `处理第 ${currentCount}/${totalUrls} 个链接失败:`, error);
                await page.close();
            }
        }

        await browser.close();
        event.reply('action-complete', {
            success: successCount,
            total: totalUrls
        });
    } catch (error) {
        console.error('\x1b[31m%s\x1b[0m', '执行操作失败:', error);
        event.reply('action-error', error.message);
    }
});
