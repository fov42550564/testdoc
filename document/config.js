const config = {
    projectName: 'Documents',
    title: '四面通测试文档',
    favicon: 'img/favicon.ico',
    logo: 'img/logo.png',
    colors: {
        primaryColor: 'rgb(34, 34, 34)',
        secondaryColor: '#FF8C00',
        activeColor: '#FF4040',
        tintColor: '#005068'
    },
    highlight: {
        theme: 'solarized-dark'
    },
    documentPath: 'docs', //默认为docs
    styles: [],
    scripts: [],
    footer: 'lib/Footer.js', //设置footer
    homePage: {
        name: '四面通测试文档',
        path: 'index.md'
    },
    menus: [
        {
            name: '业务员',
            groups: [{
                pages: [
                    { name: '个人中心', path: '业务员/个人中心.md' },
                    { name: '功能', path: '业务员/功能.md' },
                ]
            }],
        },
        {
            name: '司机端',
            groups: [{
                pages: [
                    { name: '个人中心', path: '司机端/个人中心.md' },
                ]
            }],
        },
    ]
};

module.exports = config;
