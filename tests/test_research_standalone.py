import os
import sys
from modules.scientific_research import ScientificResearcher
from colorama import init, Fore, Style

# Initialize colorama
init()

def test_research_module():
    print(f"{Fore.CYAN}=== TESTING SCIENTIFIC RESEARCH MODULE ==={Style.RESET_ALL}")
    
    # Create output directory
    output_dir = os.path.join(os.getcwd(), "test_output")
    os.makedirs(output_dir, exist_ok=True)
    print(f"Test Output Directory: {output_dir}")
    
    researcher = ScientificResearcher()
    topic = "Quantum Computing"
    
    # 1. Test Paper Search
    print(f"\n{Fore.YELLOW}[1] Testing Paper Search...{Style.RESET_ALL}")
    papers = researcher.search_papers(topic, max_results=2)
    print(f"Papers found: {len(papers)}")
    
    # 2. Test Paper Save
    if papers:
        print(f"\n{Fore.YELLOW}[2] Testing Paper Save...{Style.RESET_ALL}")
        papers_dir = os.path.join(output_dir, "research_papers")
        researcher.save_research_papers(papers, papers_dir)
        print(f"Saved papers to: {papers_dir}")
        print(f"Files in paper dir: {os.listdir(papers_dir) if os.path.exists(papers_dir) else 'Dir not found'}")
    
    # 3. Test Book Search
    print(f"\n{Fore.YELLOW}[3] Testing Book Search...{Style.RESET_ALL}")
    books = researcher.search_books_metadata(topic, max_results=2)
    print(f"Books found: {len(books)}")
    
    # 4. Test Book Download & Save
    if books:
        print(f"\n{Fore.YELLOW}[4] Testing Book Download & Save (LibGen/Internet Archive)...{Style.RESET_ALL}")
        books_dir = os.path.join(output_dir, "books")
        # Try to save just the first book to save time
        researcher.save_books_to_folder(books[:1], topic, output_dir=books_dir)
        print(f"Saved books to: {books_dir}")
        print(f"Files in book dir: {os.listdir(books_dir) if os.path.exists(books_dir) else 'Dir not found'}")

if __name__ == "__main__":
    try:
        test_research_module()
        print(f"\n{Fore.GREEN}Test Complete.{Style.RESET_ALL}")
    except Exception as e:
        print(f"\n{Fore.RED}Test Failed: {e}{Style.RESET_ALL}")
        import traceback
        traceback.print_exc()
