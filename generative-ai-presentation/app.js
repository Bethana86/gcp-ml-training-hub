// Google Cloud Introduction to Generative AI - Presentation Navigation Core

document.addEventListener('DOMContentLoaded', () => {
    const slides = document.querySelectorAll('.slide');
    const progressBarFill = document.getElementById('progress-bar-fill');
    const prevBtn = document.getElementById('prev-btn');
    const nextBtn = document.getElementById('next-btn');
    const pageNumDisplay = document.getElementById('page-num');
    
    let currentSlide = 0;
    const totalSlides = slides.length;
    
    function goToSlide(index) {
        // Enforce bounds
        if (index < 0 || index >= totalSlides) return;
        
        // Update states
        currentSlide = index;
        
        // Remove classes and rebuild positions
        slides.forEach((slide, idx) => {
            slide.classList.remove('active', 'prev-slide');
            
            if (idx < currentSlide) {
                slide.classList.add('prev-slide');
            } else if (idx === currentSlide) {
                slide.classList.add('active');
            }
        });
        
        // Update progress bar
        const progressPercent = ((currentSlide + 1) / totalSlides) * 100;
        if (progressBarFill) {
            progressBarFill.style.width = `${progressPercent}%`;
        }
        
        // Update counter text
        if (pageNumDisplay) {
            pageNumDisplay.textContent = `Slide ${currentSlide + 1} of ${totalSlides}`;
        }
    }
    
    // Mouse Button Event listeners
    if (prevBtn) {
        prevBtn.addEventListener('click', () => {
            if (currentSlide > 0) goToSlide(currentSlide - 1);
        });
    }
    
    if (nextBtn) {
        nextBtn.addEventListener('click', () => {
            if (currentSlide < totalSlides - 1) goToSlide(currentSlide + 1);
        });
    }
    
    // Keyboard Event listeners
    document.addEventListener('keydown', (e) => {
        switch (e.key) {
            case 'ArrowRight':
            case ' ':
            case 'PageDown':
                e.preventDefault();
                if (currentSlide < totalSlides - 1) {
                    goToSlide(currentSlide + 1);
                }
                break;
            case 'ArrowLeft':
            case 'PageUp':
                e.preventDefault();
                if (currentSlide > 0) {
                    goToSlide(currentSlide - 1);
                }
                break;
            case 'Home':
                e.preventDefault();
                goToSlide(0);
                break;
            case 'End':
                e.preventDefault();
                goToSlide(totalSlides - 1);
                break;
        }
    });
    
    // Initial load
    goToSlide(0);
});
