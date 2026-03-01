"""
Visualization Module - Download Relevant Images from Google
"""

import os
import re
import time
import requests
from selenium.webdriver.common.by import By
import matplotlib.pyplot as plt
import matplotlib
matplotlib.use('Agg')
from colorama import Fore, Style


def download_image(url, save_path, timeout=15):
    """Download an image from URL and convert to PNG if needed"""
    try:
        # Skip invalid URLs
        if not url or url.startswith('data:') or len(url) < 10:
            return False
        
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
            'Accept': 'image/avif,image/webp,image/apng,image/svg+xml,image/*,*/*;q=0.8',
            'Accept-Language': 'en-US,en;q=0.9',
            'Referer': 'https://www.google.com/'
        }
        
        response = requests.get(url, headers=headers, timeout=timeout, stream=True, allow_redirects=True)
        response.raise_for_status()
        
        # Check if response is actually an image
        content_type = response.headers.get('content-type', '')
        if 'image' not in content_type and 'octet-stream' not in content_type:
            return False
        
        # Download to temporary file first
        temp_path = save_path + '.temp'
        with open(temp_path, 'wb') as f:
            for chunk in response.iter_content(chunk_size=8192):
                if chunk:
                    f.write(chunk)
        
        # Verify downloaded
        if not os.path.exists(temp_path) or os.path.getsize(temp_path) < 5000:
            if os.path.exists(temp_path):
                os.remove(temp_path)
            return False
        
        # Convert to PNG using Pillow (supports WEBP, JPEG, etc.)
        try:
            from PIL import Image
            img = Image.open(temp_path)
            
            # Convert to RGB if necessary
            if img.mode in ('RGBA', 'LA', 'P'):
                # Create white background for transparency
                background = Image.new('RGB', img.size, (255, 255, 255))
                if img.mode == 'P':
                    img = img.convert('RGBA')
                if img.mode in ('RGBA', 'LA'):
                    background.paste(img, mask=img.split()[-1])
                    img = background
            elif img.mode != 'RGB':
                img = img.convert('RGB')
            
            # Ensure save_path ends with .png
            if not save_path.lower().endswith('.png'):
                save_path = os.path.splitext(save_path)[0] + '.png'
            
            # Save as PNG
            img.save(save_path, 'PNG', optimize=True)
            
            # Remove temp file
            os.remove(temp_path)
            
            return True
            
        except Exception as e:
            # If conversion fails, just rename temp file
            if os.path.exists(temp_path):
                os.rename(temp_path, save_path)
            return True
            
    except requests.exceptions.RequestException as e:
        return False
    except Exception as e:
        return False


def search_and_download_images(driver, topic, output_dir, num_images=10):
    """
    Search Google Images for relevant diagrams and download them
    
    Args:
        driver: Selenium WebDriver instance
        topic: Topic to search for
        output_dir: Directory to save images
        num_images: Number of images to download
    
    Returns:
        List of downloaded image paths
    """
    print(f"{Fore.CYAN}[*] Searching Google Images for '{topic}' diagrams...{Style.RESET_ALL}")
    
    downloaded_images = []
    
    # Create output directory if it doesn't exist
    os.makedirs(output_dir, exist_ok=True)
    
    try:
        # Navigate to Google Images
        search_query = f"{topic} diagram infographic chart"
        google_images_url = f"https://www.google.com/search?tbm=isch&q={search_query.replace(' ', '+')}"
        
        driver.get(google_images_url)
        time.sleep(5)
        
        # Scroll to load more images
        for _ in range(3):
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(1)
        
        # Get page source and extract image URLs using regex
        page_source = driver.page_source
        
        # Google Images embeds actual image URLs in the page source in various formats
        # Pattern to match image URLs in Google Images page source
        # Looks for URLs that are actual images (not thumbnails)
        url_patterns = [
            r'"(https://[^"]+\.(?:jpg|jpeg|png|gif|webp))"',
            r'\"ou\":\"([^"]+)\"',  # "ou" field contains original URL
        ]
        
        found_urls = set()
        for pattern in url_patterns:
            matches = re.findall(pattern, page_source)
            for match in matches:
                # Decode URL if needed
                url = match.replace('\\u003d', '=').replace('\\u0026', '&')
                # Skip Google's own URLs and data URLs
                if url and not url.startswith('data:') and 'google' not in url and 'gstatic' not in url:
                    if len(url) > 20 and ('http://' in url or 'https://' in url):
                        found_urls.add(url)
        
        print(f"{Fore.GREEN}[+] Found {len(found_urls)} potential image URLs{Style.RESET_ALL}")
        
        # Download images
        downloaded_count = 0
        for idx, img_url in enumerate(list(found_urls)):
            if downloaded_count >= num_images:
                break
            
            if idx >= 30:  # Try max 30 URLs
                break
            
            try:
                image_path = os.path.join(output_dir, f"diagram_{downloaded_count + 1}.png")
                
                print(f"{Fore.CYAN}   Attempting download {downloaded_count + 1}...{Style.RESET_ALL}")
                
                if download_image(img_url, image_path):
                    downloaded_images.append(image_path)
                    downloaded_count += 1
                    print(f"{Fore.GREEN}   [+] Downloaded: diagram_{downloaded_count}.png{Style.RESET_ALL}")
                
                time.sleep(0.5)
                
            except Exception as e:
                continue
        
        if downloaded_images:
            print(f"{Fore.GREEN}[+] Successfully downloaded {len(downloaded_images)} images{Style.RESET_ALL}")
        else:
            print(f"{Fore.YELLOW}[!] No images downloaded, will create sample charts{Style.RESET_ALL}")
        
        return downloaded_images
        
    except Exception as e:
        print(f"{Fore.RED}[!] Error searching images: {e}{Style.RESET_ALL}")
        return []


def create_sample_visualizations(output_dir):
    """Create sample charts as fallback"""
    print(f"{Fore.CYAN}[*] Creating sample visualizations...{Style.RESET_ALL}")
    
    os.makedirs(output_dir, exist_ok=True)
    chart_paths = []
    
    try:
        # Chart 1: Bar chart
        fig, ax = plt.subplots(figsize=(10, 6))
        categories = ['Category A', 'Category B', 'Category C', 'Category D', 'Category E']
        values = [65, 45, 78, 52, 88]
        colors = ['#3498db', '#e74c3c', '#2ecc71', '#f39c12', '#9b59b6']
        
        ax.bar(categories, values, color=colors)
        ax.set_ylabel('Value', fontsize=12)
        ax.set_title('Sample Data Distribution', fontsize=14, fontweight='bold')
        ax.grid(axis='y', alpha=0.3)
        
        chart_path = os.path.join(output_dir, 'sample_chart_1.png')
        plt.tight_layout()
        plt.savefig(chart_path, dpi=150, bbox_inches='tight')
        plt.close()
        chart_paths.append(chart_path)
        
        # Chart 2: Pie chart
        fig, ax = plt.subplots(figsize=(8, 8))
        sizes = [30, 25, 20, 15, 10]
        labels = ['Segment 1', 'Segment 2', 'Segment 3', 'Segment 4', 'Segment 5']
        colors = ['#ff6b6b', '#4ecdc4', '#45b7d1', '#f9ca24', '#6c5ce7']
        
        ax.pie(sizes, labels=labels, colors=colors, autopct='%1.1f%%', startangle=90)
        ax.set_title('Sample Distribution', fontsize=14, fontweight='bold')
        
        chart_path = os.path.join(output_dir, 'sample_chart_2.png')
        plt.savefig(chart_path, dpi=150, bbox_inches='tight')
        plt.close()
        chart_paths.append(chart_path)
        
        # Chart 3: Line chart
        fig, ax = plt.subplots(figsize=(10, 6))
        years = ['2020', '2021', '2022', '2023', '2024']
        values = [20, 35, 48, 62, 78]
        
        ax.plot(years, values, marker='o', linewidth=2, markersize=8, color='#3498db')
        ax.fill_between(range(len(years)), values, alpha=0.3, color='#3498db')
        ax.set_xlabel('Year', fontsize=12)
        ax.set_ylabel('Growth', fontsize=12)
        ax.set_title('Sample Growth Trend', fontsize=14, fontweight='bold')
        ax.grid(True, alpha=0.3)
        
        chart_path = os.path.join(output_dir, 'sample_chart_3.png')
        plt.tight_layout()
        plt.savefig(chart_path, dpi=150, bbox_inches='tight')
        plt.close()
        chart_paths.append(chart_path)
        
        print(f"{Fore.GREEN}[+] Created {len(chart_paths)} sample charts{Style.RESET_ALL}")
        
    except Exception as e:
        print( f"{Fore.RED}[!] Error creating sample charts: {e}{Style.RESET_ALL}")
    
    return chart_paths


def generate_visualizations(topic, researcher, output_dir):
    """
    Main function to generate visualizations for presentation
    Downloads relevant images from Google Images
    Falls back to sample charts if needed
    """
    print(f"{Fore.CYAN}{'='*60}{Style.RESET_ALL}")
    print(f"{Fore.CYAN}VISUALIZATION GENERATION{Style.RESET_ALL}")
    print(f"{Fore.CYAN}{'='*60}{Style.RESET_ALL}")
    
    visualization_paths = []
    
    # Try to download images from Google
    if researcher and researcher.driver:
        try:
            downloaded_images = search_and_download_images(
                researcher.driver, 
                topic, 
                output_dir, 
                num_images=10
            )
            visualization_paths.extend(downloaded_images)
        except Exception as e:
            print(f"{Fore.YELLOW}[!] Image download failed: {e}{Style.RESET_ALL}")
    
    # If we don't have enough visualizations, create sample charts
    if len(visualization_paths) < 3:
        print(f"{Fore.YELLOW}[!] Not enough downloaded images, creating sample charts...{Style.RESET_ALL}")
        sample_charts = create_sample_visualizations(output_dir)
        visualization_paths.extend(sample_charts)
    
    print(f"{Fore.GREEN}[+] Total visualizations available: {len(visualization_paths)}{Style.RESET_ALL}")
    
    return visualization_paths
