const puppeteer = require('puppeteer');
const fs = require('fs').promises;
const path = require('path');

// 添加 cookie 相关的工具函数
async function saveCookies(page, configPath) {
    try {
        const cookies = await page.cookies();
        const cookieData = {
            cookies: cookies,
            timestamp: new Date().toISOString()
        };
        
        // 使用绝对路径
        const absolutePath = path.resolve(configPath);
        await fs.writeFile(absolutePath, JSON.stringify(cookieData, null, 2), 'utf8');
        console.log('Cookie 已保存到:', absolutePath);
        return true;
    } catch (error) {
        console.error('保存 Cookie 失败:', error);
        return false;
    }
}

async function clearCookies(page) {
    try {
        const client = await page.target().createCDPSession();
        await client.send('Network.clearBrowserCookies');
        await client.send('Network.clearBrowserCache');
        console.log('Cookie 已清除');
        return true;
    } catch (error) {
        console.error('清除 Cookie 失败:', error);
        return false;
    }
}

// 优化延迟函数，缩短等待时间
const sleep = (ms) => new Promise(resolve => setTimeout(resolve, ms));
const randomSleep = async (min, max) => {
    const delay = Math.floor(Math.random() * (max - min + 1)) + min;
    await sleep(delay);
};

// 优化滚动函数，加快速度
async function simulateScroll(page) {
    await page.evaluate(async () => {
        const scrollHeight = document.documentElement.scrollHeight;
        let currentPosition = 0;
        const scrollStep = 500; // 增加滚动步长
        
        while (currentPosition < scrollHeight) {
            window.scrollBy(0, scrollStep);
            currentPosition += scrollStep;
            // 减少滚动间隔
            await new Promise(resolve => setTimeout(resolve, 50));
        }
    });
}

async function saveToFile(data, filename) {
    try {
        await fs.writeFile(filename, JSON.stringify(data, null, 2), 'utf8');
        console.log(`数据已保存到 ${filename}`);
    } catch (error) {
        console.error('保存文件失败：', error);
    }
}

// 在文件开头添加新的函数
async function loadCookiesFromFile(configPath) {
    try {
        const absolutePath = path.resolve(configPath);
        const data = await fs.readFile(absolutePath, 'utf8');
        const cookieData = JSON.parse(data);
        return cookieData.cookies || [];
    } catch (error) {
        console.error('加载 Cookie 失败:', error);
        return [];
    }
}

// 添加随机选择 cookie 的函数
function getRandomCookies(cookiesList) {
    if (!cookiesList || cookiesList.length === 0) return null;
    const randomIndex = Math.floor(Math.random() * cookiesList.length);
    return cookiesList[randomIndex];
}

// 修改 scrapeXiaohongshu 函数签名，添加 sortType 参数
async function scrapeXiaohongshu(keyword, onNewArticle, onProgress, sortType = 'general', cookieIndex = 0) {
    let browser;
    try {
        browser = await puppeteer.launch({
            headless: false,
            defaultViewport: null,
            args: [
                '--window-size=1920,1080',
                '--disable-web-security',
                '--disable-features=IsolateOrigins,site-per-process'
            ]
        });

        const page = await browser.newPage();

        // 设置 cookie
        try {
            const configPath = path.join(require('electron').app.getPath('userData'), 'config.js');
            const configContent = await fs.readFile(configPath, 'utf8');
            const config = JSON.parse(configContent);

            if (config.cookiesList && config.cookiesList.length > 0) {
                // 使用指定的 cookie
                const selectedCookies = config.cookiesList[cookieIndex];
                const cookieNote = config.cookieNotes[cookieIndex] || `Cookie组 ${cookieIndex + 1}`;
                
                if (selectedCookies) {
                    await page.setCookie(...selectedCookies);
                    console.log(`使用 Cookie: ${cookieNote}`);
                    
                    if (onProgress) {
                        onProgress({
                            currentPage: 0,
                            totalItems: 0,
                            cookieNote: cookieNote
                        });
                    }
                }
            }
        } catch (cookieError) {
            console.error('设置 Cookie 时出错:', cookieError);
        }

        // 设置请求头
        await page.setExtraHTTPHeaders({
            'Accept-Language': 'zh-CN,zh;q=0.9',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8',
            'Accept-Encoding': 'gzip, deflate, br'
        });

        await page.setUserAgent('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36');

        // 设置更短的超时时间
        page.setDefaultTimeout(5000);

        // 构建基础 URL - 不添加排序参数
        const encodedKeyword = encodeURIComponent(keyword);
        const searchUrl = `https://www.xiaohongshu.com/search_result/?keyword=${encodedKeyword}&source=web_search_result_notes`;
        
        // 访问搜索页面
        await page.goto(searchUrl, {
            waitUntil: 'networkidle0',
            timeout: 30000
        });

        // 等待页面加载完成
        await sleep(2000);

        // 修改筛选逻辑
        await page.evaluate((type) => {
            try {
                // 获取筛选下拉容器
                const dropdownContainers = document.querySelectorAll('.dropdown-container');
                if (dropdownContainers.length >= 3) {
                    // 获取第三个下拉容器
                    const targetDropdown = dropdownContainers[2];
                    if (targetDropdown) {
                        // 先点击下拉按钮
                        targetDropdown.click();
                        console.log('点击下拉按钮');

                        // 等待下拉菜单出现
                        setTimeout(() => {
                            // 获取下拉选项
                            const items = document.querySelectorAll('.dropdown-items .dropdown-item');
                            let targetIndex;
                            
                            switch(type) {
                                case 'general':
                                    targetIndex = 0;  // 综合
                                    break;
                                case 'newest':
                                    targetIndex = 1;  // 最新
                                    break;
                                case 'hot':
                                    targetIndex = 2;  // 最热
                                    break;
                                default:
                                    return;
                            }

                            // 点击对应选项
                            if (items[targetIndex]) {
                                items[targetIndex].click();
                                console.log('点击选项:', items[targetIndex].textContent);
                            }
                        }, 500);  // 等待下拉菜单出现
                    }
                }
            } catch (error) {
                console.error('筛选操作失败:', error);
            }
        }, sortType);

        // 等待筛选生效
        await sleep(3000);

        // 等待页面重新加载和筛选生效
        await sleep(2000);
        await page.waitForSelector('.note-item', { timeout: 10000 });

        // 等待搜索结果加载
        await page.waitForSelector('.note-item', { timeout: 30000 });
        
        // 先等待一下确保内容加载完成
        await sleep(2000);

        const articles = new Set();
        let currentPage = 0;

        for (let i = 0; i < 5; i++) {
            currentPage = i + 1;
            
            // 通知进度，如果返回 false 则停止抓取
            if (onProgress && !onProgress({
                currentPage: currentPage,
                totalItems: articles.size
            })) {
                console.log('收到停止信号，停止抓取');
                break;
            }

            // 先获取当前可见的文章
            const currentArticles = await page.evaluate(() => {
                const items = document.querySelectorAll('.note-item');
                return Array.from(items).map(item => {
                    try {
                        // 获取链接元素
                        const link = item.querySelector('a');
                        let url = '';
                        if (link) {
                            let href = link.getAttribute('href');
                            
                            // 提取笔记 ID
                            let noteId = '';
                            if (href.includes('/explore/')) {
                                noteId = href.match(/\/explore\/(\w+)/)?.[1];
                            } else if (href.includes('/discovery/item/')) {
                                noteId = href.match(/\/discovery\/item\/(\w+)/)?.[1];
                            } else {
                                noteId = href.match(/\/(\w+)(?:\?|$)/)?.[1];
                            }

                            // 使用标准的小红书笔记链接格式
                            url = noteId ? `https://www.xiaohongshu.com/explore/${noteId}` : '';
                        }

                        // 修改作者数据获取方式 - 使用更精确的选择器
                        const authorElement = item.querySelector('.author-wrapper a') || 
                                            item.querySelector('.user-info-wrapper a') ||
                                            item.querySelector('.user-name');
                        const author = {
                            name: '',
                            url: ''
                        };
                        
                        if (authorElement) {
                            author.name = authorElement.textContent.trim();
                            const href = authorElement.getAttribute('href');
                            if (href) {
                                // 确保链接格式正确
                                author.url = href.startsWith('/user/profile/') 
                                    ? `https://www.xiaohongshu.com${href}`
                                    : href;
                            }
                        }

                        // 调试输出
                        console.log('作者信息:', {
                            element: authorElement ? authorElement.outerHTML : 'not found',
                            name: author.name,
                            url: author.url
                        });

                        // 修改点赞数获取方式 - 使用 .count 选择器
                        let likes = 0;
                        const likeElement = item.querySelector('.count');
                        if (likeElement) {
                            const likeText = likeElement.textContent.trim();
                            // 处理可能的"万"单位
                            if (likeText.includes('万')) {
                                likes = Math.round(parseFloat(likeText.replace('万', '')) * 10000);
                            } else {
                                likes = parseInt(likeText.replace(/[^\d]/g, ''), 10);
                            }
                        }

                        const title = item.querySelector('.title')?.textContent?.trim() || '';
                        const desc = item.querySelector('.desc')?.textContent?.trim() || '';

                        if (!title && !desc) return null;

                        return {
                            title,
                            description: desc,
                            author: author.name || '未知作者',
                            authorUrl: author.url,
                            url: url,
                            likes: likes || 0,
                            timestamp: new Date().toISOString()
                        };
                    } catch (e) {
                        console.error('解析文章出错:', e);
                        return null;
                    }
                }).filter(item => item && (item.title || item.description) && item.url);
            });

            // 处理当前页面的文章
            for (const article of currentArticles) {
                if (!Array.from(articles).some(a => a.url === article.url)) {
                    articles.add(article);
                    if (onNewArticle) {
                        // 如果回调返回 false，则停止抓取
                        if (!onNewArticle(article)) {
                            console.log('收到停止信号，停止抓取');
                            return Array.from(articles);
                        }
                        await sleep(50);
                    }
                }
            }

            // 滚动页面加载更多
            await simulateScroll(page);
            // 等待新内容加载
            await sleep(2000);

            // 检查是否需要继续
            if (currentArticles.length === 0) {
                console.log('没有新数据，停止滚动');
                break;
            }
        }

        await sleep(1000);
        return Array.from(articles);
    } catch (error) {
        console.error('抓取错误：', error);
        throw error;
    } finally {
        if (browser) {
            await browser.close();
        }
    }
}

// 导出更多函数
module.exports = { 
    scrapeXiaohongshu,
    saveCookies,
    clearCookies,
    loadCookiesFromFile
};
