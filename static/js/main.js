class SlideGenerator {
    constructor() {
        this.generateBtn = document.getElementById('generate-btn');
        this.downloadBtn = document.getElementById('download-btn');
        this.topicInput = document.getElementById('topic');
        this.aiModelSelect = document.getElementById('ai-model');
        this.templateSelect = document.getElementById('template');
        this.fileInput = document.getElementById('file-input');
        this.fileStatus = document.getElementById('file-status');
        this.slidePreview = document.getElementById('slide-preview');
         this.clearFileBtn = document.getElementById('clear-file-btn');
        this.slidesData = [];
        this.currentSlideIndex = 0;
        this.slidePreview.srcdoc = this.createSlidePreviewHtml("");
        this.initializeEventListeners();
    }

    initializeEventListeners() {
        this.generateBtn.addEventListener('click', this.generateSlides.bind(this));
        this.downloadBtn.addEventListener('click', this.downloadPptx.bind(this));
        this.fileInput.addEventListener('change', this.updateFileStatus.bind(this));
        this.clearFileBtn.addEventListener('click', this.clearFileInput.bind(this));
    }

    updateFileStatus() {
        if (this.fileInput.files.length > 0) {
            this.fileStatus.textContent = this.fileInput.files[0].name;
        } else {
            this.fileStatus.textContent = 'No file selected.';
        }
    }
    clearFileInput() {
        this.fileInput.value = ''; 
        this.updateFileStatus();   
    }

    async generateSlides() {
        const topic = this.topicInput.value;
        const aiModel = this.aiModelSelect.value;
        const file = this.fileInput.files[0];
        
        if (!topic) {
            alert('Please enter a topic.');
            return;
        }

        this.showLoading(true);

        const formData = new FormData();
        formData.append('topic', topic);
        formData.append('aiModel', aiModel);
        if (file) {
            formData.append('file', file);
        }

        try {
            const response = await fetch('/api/generate', {
                method: 'POST',
                body: formData,
            });

            const responseText = await response.text();
            console.log('Raw server response:', responseText);
            const data = JSON.parse(responseText);

            if (response.ok) {
                this.slidesData = data.slides;
                let slideHtml = '';
                
                for (let i = 0; i < this.slidesData.length; i++) {
                    const slide = this.slidesData[i];
                    let imageUrl = '';
                    let imageQuery = slide.image_query || topic;

                    // Skip the Unsplash search if we already have an imagePath from the server
                    if (!slide.imagePath && imageQuery) {
                        try {
                            const imageResponse = await fetch(`/api/image?q=${encodeURIComponent(imageQuery)}`);
                            if (imageResponse.ok) {
                                const imageData = await imageResponse.json();
                                imageUrl = imageData.imageUrl;
                                slide.imageUrl = imageUrl;
                            }
                        } catch (error) {
                            console.error('Error fetching image:', error);
                        }
                    }
                    
                    slideHtml += this.createModernSlideHtml(slide, i, imageUrl);
                }

                this.slidePreview.srcdoc = this.createSlidePreviewHtml(slideHtml);
            } else {
                alert(`Error: ${data.error}`);
            }
        } catch (error) {
            console.error('Error generating slides:', error);
            alert('An error occurred while generating the slides.');
        } finally {
            this.showLoading(false);
        }
    }

createModernSlideHtml(slide, index) {
        const template = this.templateSelect.value;
        const templateConfig = this.getTemplateConfig(template);

        // ✅ NEW LOGIC: Check for the 'images' property in the slide object
        // This handles the new multi-image slide format from the server
        if (slide.images && slide.images.length > 0) {
            return this.createMultiImageSlideHtml(slide, index);
        }
        
        // Check for the old 'imagePath' for single-image slides from documents
        if (slide.imagePath) {
            return this.createImageSlideHtml(slide, index);
        }
        
        // This is the old logic for slides with HTML content
        const slideType = this.identifySlideType(slide.html, index);
        
        return `
            <div class="modern-slide" data-slide-index="${index}" style="
                background: ${templateConfig.background};
                border: none;
                border-radius: 16px;
                box-shadow: 0 20px 40px rgba(0,0,0,0.1);
                margin-bottom: 40px;
                padding: 0;
                height: 540px;
                width: 960px;
                overflow: hidden;
                position: relative;
                font-family: 'Inter', 'Segoe UI', system-ui, sans-serif;
            ">
                ${this.createSlideLayout(slide, slideType, templateConfig, slide.imageUrl)}
                
                <div style="
                    position: absolute;
                    top: -50px;
                    right: -50px;
                    width: 200px;
                    height: 200px;
                    background: ${templateConfig.accentColor};
                    border-radius: 50%;
                    opacity: 0.05;
                "></div>
                
                <div style="
                    position: absolute;
                    bottom: -30px;
                    left: -30px;
                    width: 120px;
                    height: 120px;
                    background: ${templateConfig.secondaryColor};
                    border-radius: 50%;
                    opacity: 0.08;
                "></div>
            </div>
        `;
    }

    // ✅ NEW FUNCTION: To render slides with multiple extracted images
    createMultiImageSlideHtml(slide, index) {
        const template = this.templateSelect.value;
        const templateConfig = this.getTemplateConfig(template);

        return `
            <div class="modern-slide" data-slide-index="${index}" style="
                background: ${templateConfig.background};
                border: none;
                border-radius: 16px;
                box-shadow: 0 20px 40px rgba(0,0,0,0.1);
                margin-bottom: 40px;
                padding: 0;
                height: 540px;
                width: 960px;
                overflow: hidden;
                position: relative;
                font-family: 'Inter', 'Segoe UI', system-ui, sans-serif;
            ">
                <div style="
                    height: 100%;
                    display: flex;
                    flex-direction: column;
                    justify-content: center;
                    align-items: center;
                    padding: 50px 60px;
                ">
                    <h2 style="
                        font-size: 2.5rem;
                        font-weight: 700;
                        color: ${templateConfig.primaryText};
                        margin-bottom: 30px;
                        line-height: 1.2;
                        position: relative;
                        text-align: center;
                    ">
                        ${slide.title}
                        <div style="
                            position: absolute;
                            bottom: -10px;
                            left: 50%;
                            transform: translateX(-50%);
                            width: 80px;
                            height: 4px;
                            background: ${templateConfig.accentColor};
                            border-radius: 2px;
                        "></div>
                    </h2>
                    <div style="
                        display: grid;
                        grid-template-columns: repeat(auto-fit, minmax(400px, 1fr));
                        gap: 20px;
                        width: 100%;
                        max-height: 400px;
                        overflow: hidden;
                        justify-items: center;
                        align-items: center;
                    ">
                        ${slide.images.map((img, i) => `
                            <img src="${img.url}" style="
                                width: 100%;
                                height: auto;
                                max-height: ${img.height > img.width ? '200px' : '150px'};
                                object-fit: contain;
                                border-radius: 12px;
                                box-shadow: 0 10px 20px rgba(0,0,0,0.15);
                            " alt="Image ${i + 1}" />
                        `).join('')}
                    </div>
                </div>
            </div>
        `;
    }

    // ✅ NEW FUNCTION: To render slides with extracted images
    createImageSlideHtml(slide, index) {
        const template = this.templateSelect.value;
        const templateConfig = this.getTemplateConfig(template);

        return `
            <div class="modern-slide" data-slide-index="${index}" style="
                background: ${templateConfig.background};
                border: none;
                border-radius: 16px;
                box-shadow: 0 20px 40px rgba(0,0,0,0.1);
                margin-bottom: 40px;
                padding: 0;
                height: 540px;
                width: 960px;
                overflow: hidden;
                position: relative;
                font-family: 'Inter', 'Segoe UI', system-ui, sans-serif;
            ">
                <div style="
                    height: 100%;
                    display: flex;
                    flex-direction: column;
                    justify-content: center;
                    align-items: center;
                    padding: 50px 60px;
                ">
                    <h2 style="
                        font-size: 2.5rem;
                        font-weight: 700;
                        color: ${templateConfig.primaryText};
                        margin-bottom: 30px;
                        line-height: 1.2;
                        position: relative;
                        text-align: center;
                    ">
                        ${slide.title}
                        <div style="
                            position: absolute;
                            bottom: -10px;
                            left: 50%;
                            transform: translateX(-50%);
                            width: 80px;
                            height: 4px;
                            background: ${templateConfig.accentColor};
                            border-radius: 2px;
                        "></div>
                    </h2>
                    <img src="${slide.imagePath}" style="
                        max-width: 80%;
                        max-height: 400px;
                        object-fit: contain;
                        border-radius: 16px;
                        box-shadow: 0 25px 50px rgba(0,0,0,0.2);
                    " alt="${slide.title}" />
                </div>
            </div>
        `;
    }

    createSlideLayout(slide, slideType, templateConfig, imageUrl) {
        const tempDiv = document.createElement('div');
        tempDiv.innerHTML = slide.html;
        
        const title = tempDiv.querySelector('h1, h2')?.textContent || '';
        const content = Array.from(tempDiv.querySelectorAll('p, li')).map(el => el.textContent);
        
        if (slide.table) {
            return this.createTableSlide(title, slide.table, templateConfig, imageUrl);
        }

        if (slide.chart) {
            return this.createChartSlide(title, slide.chart, templateConfig, imageUrl);
        }

        switch (slideType) {
            case 'title':
                return this.createTitleSlide(title, content, templateConfig, imageUrl);
            case 'content':
                return this.createContentSlide(title, content, templateConfig, imageUrl);
            case 'conclusion':
                return this.createConclusionSlide(title, content, templateConfig, imageUrl);
            default:
                return this.createContentSlide(title, content, templateConfig, imageUrl);
        }
    }

    createTitleSlide(title, content, templateConfig, imageUrl) {
        return `
            <div style="
                height: 100%;
                display: flex;
                flex-direction: column;
                justify-content: center;
                align-items: center;
                text-align: center;
                padding: 60px;
                background: ${imageUrl ? `linear-gradient(rgba(0,0,0,0.4), rgba(0,0,0,0.4)), url('${imageUrl}')` : templateConfig.gradient};
                background-size: cover;
                background-position: center;
                color: ${templateConfig.primaryText};
            ">
                <div style="
                    background: ${imageUrl ? 'rgba(255,255,255,0.95)' : 'transparent'};
                    padding: ${imageUrl ? '40px' : '0'};
                    border-radius: ${imageUrl ? '20px' : '0'};
                    backdrop-filter: ${imageUrl ? 'blur(10px)' : 'none'};
                    color: ${imageUrl ? templateConfig.primaryText : 'inherit'};
                ">
                    <h1 style="
                        font-size: 3.5rem;
                        font-weight: 800;
                        margin-bottom: 24px;
                        line-height: 1.1;
                        color: ${templateConfig.primaryText};
                    ">${title}</h1>
                    ${content.length > 0 ? `
                        <p style="
                            font-size: 1.4rem;
                            opacity: 0.8;
                            max-width: 600px;
                            line-height: 1.6;
                            margin: 0 auto;
                            color: ${templateConfig.secondaryText};
                        ">${content[0]}</p>
                    ` : ''}
                </div>
            </div>
        `;
    }

    createContentSlide(title, content, templateConfig, imageUrl) {
        return `
            <div style="
                height: 100%;
                display: grid;
                grid-template-columns: ${imageUrl ? '1fr 1fr' : '1fr'};
                gap: 40px;
                padding: 50px 60px;
            ">
                <div style="
                    display: flex;
                    flex-direction: column;
                    justify-content: center;
                ">
                    <h2 style="
                        font-size: 2.5rem;
                        font-weight: 700;
                        color: ${templateConfig.primaryText};
                        margin-bottom: 30px;
                        line-height: 1.2;
                        position: relative;
                    ">
                        ${title}
                        <div style="
                            position: absolute;
                            bottom: -10px;
                            left: 0;
                            width: 80px;
                            height: 4px;
                            background: ${templateConfig.accentColor};
                            border-radius: 2px;
                        "></div>
                    </h2>
                    
                    <div style="space-y: 16px;">
                        ${content.map((text, index) => `
                            <div style="
                                display: flex;
                                align-items: flex-start;
                                margin-bottom: 20px;
                                animation: slideInUp 0.6s ease-out ${index * 0.1}s both;
                            ">
                                <div style="
                                    width: 8px;
                                    height: 8px;
                                    background: ${templateConfig.accentColor};
                                    border-radius: 50%;
                                    margin-top: 8px;
                                    margin-right: 16px;
                                    flex-shrink: 0;
                                "></div>
                                <p style="
                                    font-size: 1.1rem;
                                    line-height: 1.6;
                                    color: ${templateConfig.secondaryText};
                                    margin: 0;
                                ">${text}</p>
                            </div>
                        `).join('')}
                    </div>
                </div>
                
                ${imageUrl ? `
                    <div style="
                        display: flex;
                        align-items: center;
                        justify-content: center;
                        position: relative;
                    ">
                        <div style="
                            position: relative;
                            border-radius: 16px;
                            overflow: hidden;
                            box-shadow: 0 25px 50px rgba(0,0,0,0.2);
                            transform: rotate(-2deg);
                            transition: transform 0.3s ease;
                        ">
                            <img src="${imageUrl}" style="
                                width: 100%;
                                height: 300px;
                                object-fit: cover;
                                display: block;
                            " alt="Slide image" />
                            <div style="
                                position: absolute;
                                inset: 0;
                                background: linear-gradient(45deg, transparent 0%, rgba(255,255,255,0.1) 100%);
                            "></div>
                        </div>
                    </div>
                ` : ''}
            </div>
        `;
    }

    createTableSlide(title, table, templateConfig, imageUrl) {
        return `
            <div style="
                height: 100%;
                display: flex;
                flex-direction: column;
                justify-content: center;
                padding: 50px 60px;
            ">
                <h2 style="
                    font-size: 2.5rem;
                    font-weight: 700;
                    color: ${templateConfig.primaryText};
                    margin-bottom: 30px;
                    line-height: 1.2;
                    position: relative;
                    text-align: center;
                ">
                    ${title}
                    <div style="
                        position: absolute;
                        bottom: -10px;
                        left: 50%;
                        transform: translateX(-50%);
                        width: 80px;
                        height: 4px;
                        background: ${templateConfig.accentColor};
                        border-radius: 2px;
                    "></div>
                </h2>
                <div style="overflow-x: auto;">
                    <table style="width: 100%; border-collapse: collapse;">
                        <thead>
                            <tr>
                                ${table.headers.map(header => `<th style="padding: 12px 15px; text-align: left; background-color: ${templateConfig.accentColor}; color: white; font-weight: 600;">${header}</th>`).join('')}
                            </tr>
                        </thead>
                        <tbody>
                            ${table.rows.map(row => `
                                <tr>
                                    ${row.map(cell => `<td style="padding: 12px 15px; border-bottom: 1px solid #e2e8f0; color: ${templateConfig.secondaryText};">${cell}</td>`).join('')}
                                </tr>
                            `).join('')}
                        </tbody>
                    </table>
                </div>
            </div>
        `;
    }

    createChartSlide(title, chart, templateConfig, imageUrl) {
        const chartId = `chart-${Date.now()}`;
        setTimeout(() => this.renderChart(chartId, chart), 100);
        return `
            <div style="
                height: 100%;
                display: flex;
                flex-direction: column;
                justify-content: center;
                padding: 50px 60px;
            ">
                <h2 style="
                    font-size: 2.5rem;
                    font-weight: 700;
                    color: ${templateConfig.primaryText};
                    margin-bottom: 30px;
                    line-height: 1.2;
                    position: relative;
                    text-align: center;
                ">
                    ${title}
                    <div style="
                        position: absolute;
                        bottom: -10px;
                        left: 50%;
                        transform: translateX(-50%);
                        width: 80px;
                        height: 4px;
                        background: ${templateConfig.accentColor};
                        border-radius: 2px;
                    "></div>
                </h2>
                <div id="${chartId}" style="width: 100%; height: 400px;"></div>
            </div>
        `;
    }

    renderChart(chartId, chart) {
        const chartDom = this.slidePreview.contentWindow.document.getElementById(chartId);
        if (chartDom) {
            const myChart = echarts.init(chartDom);
            const option = {
                xAxis: {
                    type: 'category',
                    data: chart.data.labels
                },
                yAxis: {
                    type: 'value'
                },
                series: chart.data.datasets.map(dataset => ({
                    data: dataset.data,
                    type: chart.type || 'bar'
                }))
            };
            myChart.setOption(option);
        }
    }

    createConclusionSlide(title, content, templateConfig, imageUrl) {
        return `
            <div style="
                height: 100%;
                display: flex;
                flex-direction: column;
                justify-content: center;
                align-items: center;
                text-align: center;
                padding: 60px;
                background: ${templateConfig.gradient};
                position: relative;
            ">
                <div style="
                    position: absolute;
                    inset: 0;
                    background: url('data:image/svg+xml,<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100"><defs><pattern id="grain" width="100" height="100" patternUnits="userSpaceOnUse"><circle cx="20" cy="20" r="1" fill="white" opacity="0.1"/><circle cx="80" cy="40" r="1" fill="white" opacity="0.1"/><circle cx="40" cy="80" r="1" fill="white" opacity="0.1"/></pattern></defs><rect width="100" height="100" fill="url(%23grain)"/></svg>');
                "></div>
                
                <div style="position: relative; z-index: 1;">
                    <h2 style="
                        font-size: 2rem;
                        font-weight: 700;
                        color: ${templateConfig.primaryText};
                        margin-bottom: 30px;
                        text-shadow: 0 2px 4px rgba(0,0,0,0.3);
                    ">${title}</h2>
                    
                    ${content.map(text => `
                        <p style="
                            font-size: 1.3rem;
                            color: ${templateConfig.secondaryText};
                            max-width: 700px;
                            line-height: 1.6;
                            margin-bottom: 20px;
                        ">${text}</p>
                    `).join('')}
                    
                    <div style="
                        margin-top: 40px;
                        padding: 16px 32px;
                        background: rgba(255,255,255,0.2);
                        border: 2px solid rgba(255,255,255,0.3);
                        border-radius: 12px;
                        backdrop-filter: blur(10px);
                        display: inline-block;
                        font-weight: 600;
                        color: white;
                        font-size: 1.1rem;
                    ">
                        Thank You!
                    </div>
                </div>
            </div>
        `;
    }

    identifySlideType(html, index) {
        const lowerHtml = html.toLowerCase();
        if (index === 0 || lowerHtml.includes('title') || lowerHtml.includes('<h1')) {
            return 'title';
        }
        if (lowerHtml.includes('conclusion') || lowerHtml.includes('thank') || lowerHtml.includes('summary')) {
            return 'conclusion';
        }
        return 'content';
    }

    getTemplateConfig(template) {
        const templates = {
            'Default': {
                background: 'linear-gradient(135deg, #667eea 0%, #764ba2 100%)',
                gradient: 'linear-gradient(135deg, #667eea 0%, #764ba2 100%)',
                textGradient: 'linear-gradient(135deg, #667eea 0%, #764ba2 100%)',
                primaryText: '#FFFFFF',
                secondaryText: '#E2E8F0',
                accentColor: '#667eea',
                secondaryColor: '#764ba2'
            },
            'Dark': {
                background: 'linear-gradient(135deg, #1a202c 0%, #2d3748 100%)',
                gradient: 'linear-gradient(135deg, #1a202c 0%, #2d3748 100%)',
                textGradient: 'linear-gradient(135deg, #63b3ed 0%, #90cdf4 100%)',
                primaryText: '#f7fafc',
                secondaryText: '#e2e8f0',
                accentColor: '#63b3ed',
                secondaryColor: '#90cdf4'
            },
            'Light': {
                background: 'linear-gradient(135deg, #f7fafc 0%, #edf2f7 100%)',
                gradient: 'linear-gradient(135deg, #4299e1 0%, #3182ce 100%)',
                textGradient: 'linear-gradient(135deg, #4299e1 0%, #3182ce 100%)',
                primaryText: '#1a202c',
                secondaryText: '#2d3748',
                accentColor: '#4299e1',
                secondaryColor: '#63b3ed'
            },
            'Blue': {
                background: 'linear-gradient(135deg, #ebf8ff 0%, #bee3f8 100%)',
                gradient: 'linear-gradient(135deg, #2b6cb0 0%, #2c5282 100%)',
                textGradient: 'linear-gradient(135deg, #2b6cb0 0%, #2c5282 100%)',
                primaryText: '#1a365d',
                secondaryText: '#2c5282',
                accentColor: '#3182ce',
                secondaryColor: '#4299e1'
            }
        };
        
        return templates[template] || templates['Default'];
    }

    createSlidePreviewHtml(slideHtml) {
        return `
            <html>
                <head>
                    <link rel="stylesheet" href="/static/css/bootstrap.min.css">
                    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap" rel="stylesheet">
                    <script src="/static/js/echarts.min.js"></script>
                    <style>
                        body { 
                            background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%);
                            margin: 0;
                            padding: 20px;
                            font-family: 'Inter', sans-serif;
                        }
                        
                        .modern-slide {
                            transition: transform 0.3s ease, box-shadow 0.3s ease;
                        }
                        
                        .modern-slide:hover {
                            transform: translateY(-8px);
                            box-shadow: 0 30px 60px rgba(0,0,0,0.15) !important;
                        }
                        
                        @keyframes slideInUp {
                            from {
                                opacity: 0;
                                transform: translateY(20px);
                            }
                            to {
                                opacity: 1;
                                transform: translateY(0);
                            }
                        }
                        
                        @keyframes fadeIn {
                            from { opacity: 0; }
                            to { opacity: 1; }
                        }
                        
                        .slide-container {
                            display: flex;
                            flex-direction: column;
                            align-items: center;
                            gap: 30px;
                        }
                    </style>
                </head>
                <body>
                    <div class="slide-container">
                        ${slideHtml}
                    </div>
                </body>
            </html>
        `;
    }

async downloadPptx() {
        if (this.slidesData.length === 0) {
            alert("Please generate slides first.");
            return;
        }

        this.showLoading(true);

        const pptx = new PptxGenJS();
        pptx.layout = 'LAYOUT_16x9';

        const selectedTemplate = this.templateSelect.value;

        const pptxTemplates = {
            'Default': {
                background: '667eea',
                titleColor: 'FFFFFF',
                textColor: 'F8F9FA',
                accentColor: '764ba2'
            },
            'Dark': {
                background: '1a202c',
                titleColor: '63b3ed',
                textColor: 'e2e8f0',
                accentColor: '90cdf4'
            },
            'Light': {
                background: 'f7fafc',
                titleColor: '4299e1',
                textColor: '2d3748',
                accentColor: '3182ce'
            },
            'Blue': {
                background: 'ebf8ff',
                titleColor: '1a365d',
                textColor: '2c5282',
                accentColor: '3182ce'
            }
        };

        const template = pptxTemplates[selectedTemplate];

        pptx.defineSlideMaster({
            title: 'MASTER_SLIDE',
            background: template.background,
            objects: [],
        });

        for (let i = 0; i < this.slidesData.length; i++) {
            const slideData = this.slidesData[i];
            const slide = pptx.addSlide({ masterName: 'MASTER_SLIDE' });
            
            // Handle multi-image slides first (from documents)
            if (slideData.images && slideData.images.length > 0) {
                slide.addText(slideData.title, { 
                    x: '5%', y: '5%', w: '90%', h: '10%', 
                    fontSize: 24, bold: true, align: 'center', 
                    color: template.titleColor,
                    fontFace: 'Segoe UI'
                });

                // Calculate image layout for multiple images
                const totalImages = slideData.images.length;
                let imageX, imageY, imageW, imageH;
                
                switch(totalImages) {
                    case 1:
                        imageX = '15%'; imageY = '20%'; imageW = '70%'; imageH = '70%';
                        break;
                    case 2:
                        imageX = '5%'; imageY = '20%'; imageW = '42%'; imageH = '70%';
                        break;
                    case 3:
                        imageX = ['5%', '35%', '65%']; imageY = '25%'; imageW = '28%'; imageH = '50%';
                        break;
                    case 4:
                        imageX = ['5%', '52%', '5%', '52%']; imageY = ['20%', '20%', '55%', '55%']; imageW = '42%'; imageH = '35%';
                        break;
                    default:
                        imageX = '5%'; imageY = '20%'; imageW = '90%'; imageH = '70%';
                }

                try {
                    for (let j = 0; j < slideData.images.length; j++) {
                        const image = slideData.images[j];
                        const response = await fetch(image.url);
                        const blob = await response.blob();
                        const base64Data = await new Promise((resolve, reject) => {
                            const reader = new FileReader();
                            reader.onloadend = () => resolve(reader.result);
                            reader.onerror = reject;
                            reader.readAsDataURL(blob);
                        });

                        slide.addImage({
                            data: base64Data,
                            x: Array.isArray(imageX) ? imageX[j] : imageX,
                            y: Array.isArray(imageY) ? imageY[j] : imageY,
                            w: Array.isArray(imageW) ? imageW[j] : imageW,
                            h: Array.isArray(imageH) ? imageH[j] : imageH,
                            sizing: { type: 'contain' }
                        });
                    }
                } catch (error) {
                    console.error('Error adding images to PPTX:', error);
                }
            }
            // Handle single image slides (from documents)
            else if (slideData.imagePath) {
                 try {
                    const response = await fetch(slideData.imagePath);
                    const blob = await response.blob();
                    const imageData = await new Promise((resolve, reject) => {
                        const reader = new FileReader();
                        reader.onloadend = () => resolve(reader.result);
                        reader.onerror = reject;
                        reader.readAsDataURL(blob);
                    });

                    slide.addText(slideData.title, { 
                        x: '5%', y: '5%', w: '90%', h: '10%', 
                        fontSize: 24, bold: true, align: 'center', 
                        color: template.titleColor,
                        fontFace: 'Segoe UI'
                    });
                    slide.addImage({
                        data: imageData,
                        x: '10%', y: '20%', w: '80%', h: '70%',
                        sizing: { type: 'contain' }
                    });
                } catch (error) {
                    console.error('Error adding image to PPTX:', error);
                }
            }
            // ✅ Correctly handle AI-generated slides (which may also have a web image URL)
            else {
                const tempDiv = document.createElement('div');
                tempDiv.innerHTML = slideData.html;
                const title = tempDiv.querySelector('h1, h2')?.innerText || '';
                
                let imageUrl = slideData.imageUrl;
                let imageData = null;

                if (imageUrl) {
                    try {
                        const response = await fetch(imageUrl);
                        const blob = await response.blob();
                        imageData = await new Promise((resolve, reject) => {
                            const reader = new FileReader();
                            reader.onloadend = () => resolve(reader.result);
                            reader.onerror = reject;
                            reader.readAsDataURL(blob);
                        });
                    } catch (error) {
                        console.error('Error fetching image for PPTX:', error);
                    }
                }

                if (slideData.table) {
                    slide.addText(title, { 
                        x: '5%', y: '5%', w: '90%', h: '10%', 
                        fontSize: 24, bold: true, align: 'center', 
                        color: template.titleColor,
                        fontFace: 'Segoe UI'
                    });
                    slide.addTable(slideData.table.rows, { 
                        x: '5%', y: '20%', w: '90%', 
                        rowH: 0.5, 
                        fill: { color: 'F7F7F7' }, 
                        color: '3D3D3D',
                        border: { pt: 1, color: 'C7C7C7' }
                    });
                } else if (slideData.chart) {
                    slide.addText(title, { 
                        x: '5%', y: '5%', w: '90%', h: '10%', 
                        fontSize: 24, bold: true, align: 'center', 
                        color: template.titleColor,
                        fontFace: 'Segoe UI'
                    });

                    const chartData = slideData.chart.data.datasets.map(dataset => ({
                        name: dataset.label,
                        labels: slideData.chart.data.labels,
                        values: dataset.data
                    }));

                    slide.addChart(pptx.ChartType[slideData.chart.type], chartData, { 
                        x: '5%', y: '20%', w: '90%', h: '70%' 
                    });
                } else {
                    const slideType = slideData.html ? this.identifySlideType(slideData.html, i) : 'content';
                    
                    if (slideType === 'title') {
                        const subtitle = tempDiv.querySelector('p')?.innerText || '';
                        
                        slide.addText(title, { 
                            x: '10%', y: '35%', w: '80%', h: '20%', 
                            fontSize: 32, bold: true, align: 'center', 
                            color: template.titleColor,
                            fontFace: 'Segoe UI'
                        });
                        if (subtitle) {
                            slide.addText(subtitle, { 
                                x: '10%', y: '55%', w: '80%', h: '10%', 
                                fontSize: 18, align: 'center', 
                                color: template.textColor,
                                fontFace: 'Segoe UI'
                            });
                        }
                    } else if (slideType === 'content' || slideType === 'conclusion') {
                        const contentText = Array.from(tempDiv.querySelectorAll('p, li'))
                            .map(el => el.innerText.trim())
                            .filter(text => text.length > 0)
                            .join('\n');
                        
                        slide.addText(title, { 
                            x: '5%', y: '5%', w: '90%', h: '10%', 
                            fontSize: 24, bold: true, color: template.titleColor,
                            fontFace: 'Segoe UI'
                        });
                        slide.addText(contentText, {
                            x: '5%', y: '15%', w: '90%', h: '80%',
                            fontSize: 16, color: template.textColor,
                            fontFace: 'Segoe UI'
                        });
                    }
                }

                // ✅ Add the downloaded image to the slide if it exists
                if (imageData) {
                     slide.addImage({
                        data: imageData,
                        x: '60%', y: '15%', w: '30%', h: '70%',
                        sizing: { type: 'contain' }
                    });
                }
            }
        }

        pptx.writeFile({ fileName: `Presentation-${Date.now()}.pptx` });
        this.showLoading(false);
    }
       showLoading(isLoading) {
        const loadingSpinner = document.getElementById('loading-spinner');
        if (isLoading) {
            loadingSpinner.style.display = 'block';
            this.generateBtn.disabled = true;
            this.downloadBtn.disabled = true;
        } else {
            loadingSpinner.style.display = 'none';
            this.generateBtn.disabled = false;
            this.downloadBtn.disabled = false;
        }
    }
}


document.addEventListener('DOMContentLoaded', () => {
    new SlideGenerator();
});