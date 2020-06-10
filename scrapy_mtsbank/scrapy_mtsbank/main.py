from scrapy.cmdline import execute

def main():
    # execute(['scrapy', 'crawl', 'mtsbank', '--nolog'])
    execute(['scrapy', 'crawl', 'mtsbank'])

if __name__ == '__main__':
    main()