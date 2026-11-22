import requests
import base64
import hashlib
import time
import random
from difflib import SequenceMatcher

class InfoVerifier:
    def __init__(self, api_key, secret_key):
        self.api_key = api_key
        self.secret_key = secret_key
        self.cache = {}

    def get_access_token(self):
        # 获取百度API访问令牌
        url = "https://aip.baidubce.com/oauth/2.0/token"
        params = {
            "grant_type": "client_credentials",
            "client_id": self.api_key,
            "client_secret": self.secret_key
        }
        response = requests.post(url, params=params)
        return response.json().get("access_token")

    def search_baidu_api(self, query):
        if query in self.cache:
            return self.cache[query]
        
        token = self.get_access_token()
        url = "https://aip.baidubce.com/rpc/2.0/ai_custom/v1/websearch/custom_search"
        headers = {'Content-Type': 'application/json'}
        data = {
            "access_token": token,
            "query": query,
            "num": 5  # 返回结果数量
        }
        response = requests.post(url, json=data, headers=headers)
        results = response.json().get("result", {}).get("items", [])
        
        self.cache[query] = results
        time.sleep(random.uniform(1, 3))  # 随机延迟避免频率限制
        return results

    def calculate_similarity(self, text1, text2):
        return SequenceMatcher(None, text1, text2).ratio()

    def verify_credibility(self, field_data, sources=3):
        query = str(field_data)
        results = self.search_baidu_api(query)
        
        if len(results) < sources:
            return 0.0
        
        similarities = []
        for result in results[:sources]:
            combined_text = result.get('title', '') + ' ' + result.get('snippet', '')
            sim = self.calculate_similarity(str(field_data), combined_text)
            similarities.append(sim)
        
        avg_sim = sum(similarities) / len(similarities)
        credibility = min(1.0, avg_sim * 2)
        return credibility

# 示例使用（替换为你的API Key和Secret Key）
# verifier = InfoVerifier("your_api_key", "your_secret_key")
# credibility = verifier.verify_credibility("海外中式建筑案例")
# print(f"可信度: {credibility}")