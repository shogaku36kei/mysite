// ========================================
// わくわくギルド - フロントエンド設定ファイル（完全版）
// ========================================

// GAS Web App URL
const GAS_WEB_APP_URL = 'https://script.google.com/macros/s/AKfycbx6OdU_iOb4yyJ8xBOdLsrnW7MI35C7wGjiAGrWls3pUFSb5RHCvRl8q4NEQuP1quCavw/exec';

// サイト設定
const CONFIG = {
  siteName: 'わくわくギルド',
  adminEmail: 'guild@city.soja.okayama.jp',
  maxQuestReward: 1000000,
  minQuestReward: 1000,
  taxRate: 0.1021,
  taxRateHigh: 0.2042,
  taxThreshold: 1000000,
  version: '1.0.0'
};

// ランク設定（経験値とアイコン）
const RANKS = {
  F: { name: 'F', exp: 0, color: '#9e9e9e', icon: '🔰', title: '見習い冒険者' },
  E: { name: 'E', exp: 100, color: '#7c4dff', icon: '⚔️', title: '駆け出し冒険者' },
  D: { name: 'D', exp: 300, color: '#00b0ff', icon: '🛡️', title: '一人前の冒険者' },
  C: { name: 'C', exp: 600, color: '#00e676', icon: '⚡', title: '熟練冒険者' },
  B: { name: 'B', exp: 1000, color: '#ffd600', icon: '🌟', title: 'ベテラン冒険者' },
  A: { name: 'A', exp: 1500, color: '#ff6f00', icon: '👑', title: 'エリート冒険者' },
  S: { name: 'S', exp: 2500, color: '#ff1744', icon: '💎', title: '伝説の冒険者' }
};

// 難易度設定
const DIFFICULTIES = {
  1: { stars: '★', exp: 10, label: '簡単・初心者OK', color: '#4caf50' },
  2: { stars: '★★', exp: 20, label: 'やや簡単', color: '#8bc34a' },
  3: { stars: '★★★', exp: 30, label: '普通', color: '#ff9800' },
  4: { stars: '★★★★', exp: 50, label: 'やや難しい', color: '#ff5722' },
  5: { stars: '★★★★★', exp: 100, label: '高難度', color: '#f44336' }
};

// カテゴリ設定
const CATEGORIES = {
  '草刈り': { icon: '🌿', color: '#4caf50' },
  '清掃': { icon: '🧹', color: '#2196f3' },
  '力仕事': { icon: '💪', color: '#ff5722' },
  'IT支援': { icon: '💻', color: '#9c27b0' },
  '料理・調理': { icon: '🍳', color: '#ff9800' },
  'イベント手伝い': { icon: '🎉', color: '#e91e63' },
  'その他': { icon: '📦', color: '#607d8b' }
};

// ステータス定義
const STATUS = {
  // 審査状況
  pending: { label: '申請中', color: '#ff9800', icon: '⏳' },
  approved: { label: '承認', color: '#4caf50', icon: '✅' },
  rejected: { label: '却下', color: '#f44336', icon: '❌' },
  suspended: { label: '停止', color: '#000000', icon: '🚫' },
  
  // クエスト掲載状況
  recruiting: { label: '募集中', color: '#2196f3', icon: '📢' },
  accepted: { label: '受注済', color: '#9c27b0', icon: '🤝' },
  inProgress: { label: '進行中', color: '#00bcd4', icon: '⚙️' },
  completed: { label: '完了', color: '#4caf50', icon: '✅' },
  cancelled: { label: '中止', color: '#f44336', icon: '❌' },
  deleted: { label: '削除', color: '#9e9e9e', icon: '🗑️' }
};

// ========================================
// ユーティリティ関数
// ========================================

const Utils = {
  /**
   * 源泉徴収税額を計算
   */
  calculateTax: (amount) => {
    if (amount <= CONFIG.taxThreshold) {
      return Math.floor(amount * CONFIG.taxRate);
    } else {
      const baseTax = Math.floor(CONFIG.taxThreshold * CONFIG.taxRate);
      const excessTax = Math.floor((amount - CONFIG.taxThreshold) * CONFIG.taxRateHigh);
      return baseTax + excessTax;
    }
  },
  
  /**
   * 手取り額を計算
   */
  calculateNetAmount: (amount) => {
    return amount - Utils.calculateTax(amount);
  },
  
  /**
   * 経験値からランクを判定
   */
  getRankFromExp: (exp) => {
    const rankKeys = Object.keys(RANKS).reverse();
    for (let key of rankKeys) {
      if (exp >= RANKS[key].exp) {
        return key;
      }
    }
    return 'F';
  },
  
  /**
   * 次のランクまでの経験値を計算
   */
  getExpToNextRank: (currentExp) => {
    const currentRank = Utils.getRankFromExp(currentExp);
    const rankKeys = Object.keys(RANKS);
    const currentIndex = rankKeys.indexOf(currentRank);
    
    if (currentIndex === rankKeys.length - 1) {
      return 0; // 最高ランク
    }
    
    const nextRank = rankKeys[currentIndex + 1];
    return RANKS[nextRank].exp - currentExp;
  },
  
  /**
   * 次のランク名を取得
   */
  getNextRank: (currentExp) => {
    const currentRank = Utils.getRankFromExp(currentExp);
    const rankKeys = Object.keys(RANKS);
    const currentIndex = rankKeys.indexOf(currentRank);
    
    if (currentIndex === rankKeys.length - 1) {
      return null; // 最高ランク
    }
    
    return rankKeys[currentIndex + 1];
  },
  
  /**
   * 日付フォーマット
   */
  formatDate: (dateString) => {
    if (!dateString) return '-';
    const date = new Date(dateString);
    return `${date.getFullYear()}/${String(date.getMonth() + 1).padStart(2, '0')}/${String(date.getDate()).padStart(2, '0')}`;
  },
  
  /**
   * 日時フォーマット
   */
  formatDateTime: (dateString) => {
    if (!dateString) return '-';
    const date = new Date(dateString);
    return `${date.getFullYear()}/${String(date.getMonth() + 1).padStart(2, '0')}/${String(date.getDate()).padStart(2, '0')} ${String(date.getHours()).padStart(2, '0')}:${String(date.getMinutes()).padStart(2, '0')}`;
  },
  
  /**
   * 金額フォーマット
   */
  formatCurrency: (amount) => {
    if (amount === null || amount === undefined) return '¥0';
    return `¥${Number(amount).toLocaleString()}`;
  },
  
  /**
   * ステータスバッジHTML生成
   */
  getStatusBadge: (status) => {
    const s = STATUS[status];
    if (!s) return `<span class="status-badge">${status}</span>`;
    return `<span class="status-badge" style="background: ${s.color}; color: white; padding: 0.3rem 0.8rem; border-radius: 20px; font-size: 0.85rem;">${s.icon} ${s.label}</span>`;
  },
  
  /**
   * ランクバッジHTML生成
   */
  getRankBadge: (rank) => {
    const r = RANKS[rank];
    if (!r) return `<span class="rank-badge">${rank}</span>`;
    return `<span class="rank-badge rank-${rank}" style="background: ${r.color}; color: white; padding: 0.3rem 0.8rem; border-radius: 5px; font-weight: bold;">${r.icon} ${rank}</span>`;
  },
  
  /**
   * カテゴリバッジHTML生成
   */
  getCategoryBadge: (category) => {
    const c = CATEGORIES[category];
    if (!c) return `<span class="category-tag">${category}</span>`;
    return `<span class="category-tag" style="background: ${c.color}; color: white; padding: 0.3rem 0.8rem; border-radius: 20px; font-size: 0.85rem;">${c.icon} ${category}</span>`;
  },
  
  /**
   * 難易度表示HTML生成
   */
  getDifficultyStars: (difficulty) => {
    const d = DIFFICULTIES[difficulty];
    if (!d) return '★';
    return `<span style="color: ${d.color}; font-size: 1.2rem;">${d.stars}</span>`;
  },
  
  /**
   * ローディング表示
   */
  showLoading: (elementId) => {
    const element = document.getElementById(elementId);
    if (element) {
      element.innerHTML = '<div class="loading">🔄 読み込み中...</div>';
    }
  },
  
  /**
   * エラー表示
   */
  showError: (elementId, message) => {
    const element = document.getElementById(elementId);
    if (element) {
      element.innerHTML = `<div class="error">❌ ${message}</div>`;
    }
  },
  
  /**
   * データなし表示
   */
  showNoData: (elementId, message = 'データがありません') => {
    const element = document.getElementById(elementId);
    if (element) {
      element.innerHTML = `<div class="no-data"><p>${message}</p></div>`;
    }
  },
  
  /**
   * 成功メッセージ表示
   */
  showSuccess: (elementId, message) => {
    const element = document.getElementById(elementId);
    if (element) {
      element.innerHTML = `<div class="success-message">✅ ${message}</div>`;
      setTimeout(() => {
        element.innerHTML = '';
      }, 5000);
    }
  }
};

// ========================================
// API通信
// ========================================

const API = {
  /**
   * GETリクエスト
   */
  get: async (action, params = {}) => {
    const url = new URL(GAS_WEB_APP_URL);
    url.searchParams.append('action', action);
    Object.keys(params).forEach(key => {
      if (params[key] !== null && params[key] !== undefined) {
        url.searchParams.append(key, params[key]);
      }
    });
    
    try {
      const response = await fetch(url, {
        method: 'GET',
        mode: 'cors'
      });
      
      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }
      
      return await response.json();
    } catch (error) {
      console.error('API GET Error:', error);
      throw error;
    }
  },
  
  /**
   * POSTリクエスト
   */
  post: async (action, data = {}) => {
    try {
      const response = await fetch(GAS_WEB_APP_URL, {
        method: 'POST',
        mode: 'cors',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          action: action,
          data: data
        })
      });
      
      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }
      
      return await response.json();
    } catch (error) {
      console.error('API POST Error:', error);
      throw error;
    }
  }
};

// ========================================
// ローカルストレージ管理
// ========================================

const Storage = {
  /**
   * ログイン情報保存
   */
  saveLogin: (userType, userId, userData) => {
    localStorage.setItem('wg_userType', userType);
    localStorage.setItem('wg_userId', userId);
    localStorage.setItem('wg_userData', JSON.stringify(userData));
    localStorage.setItem('wg_loginTime', new Date().toISOString());
  },
  
  /**
   * ログイン情報取得
   */
  getLogin: () => {
    return {
      userType: localStorage.getItem('wg_userType'),
      userId: localStorage.getItem('wg_userId'),
      userData: JSON.parse(localStorage.getItem('wg_userData') || '{}'),
      loginTime: localStorage.getItem('wg_loginTime')
    };
  },
  
  /**
   * ログアウト
   */
  clearLogin: () => {
    localStorage.removeItem('wg_userType');
    localStorage.removeItem('wg_userId');
    localStorage.removeItem('wg_userData');
    localStorage.removeItem('wg_loginTime');
  },
  
  /**
   * ログイン状態確認
   */
  isLoggedIn: () => {
    return !!localStorage.getItem('wg_userId');
  },
  
  /**
   * 管理者ログイン状態確認
   */
  isAdmin: () => {
    return localStorage.getItem('wg_userType') === 'admin';
  }
};

// ========================================
// 通知システム
// ========================================

const Notification = {
  /**
   * 通知を表示
   */
  show: (message, type = 'info', duration = 5000) => {
    // 既存の通知を削除
    const existing = document.querySelector('.notification');
    if (existing) {
      existing.remove();
    }
    
    const notification = document.createElement('div');
    notification.className = `notification notification-${type}`;
    
    const icon = {
      success: '✅',
      error: '❌',
      warning: '⚠️',
      info: 'ℹ️'
    }[type] || 'ℹ️';
    
    notification.innerHTML = `
      <span>${icon} ${message}</span>
      <button onclick="this.parentElement.remove()" style="background: transparent; border: none; color: inherit; font-size: 1.5rem; cursor: pointer; padding: 0; margin-left: 1rem;">×</button>
    `;
    
    document.body.appendChild(notification);
    
    // 自動削除
    if (duration > 0) {
      setTimeout(() => {
        if (notification.parentElement) {
          notification.remove();
        }
      }, duration);
    }
  },
  
  success: (message, duration = 5000) => Notification.show(message, 'success', duration),
  error: (message, duration = 5000) => Notification.show(message, 'error', duration),
  warning: (message, duration = 5000) => Notification.show(message, 'warning', duration),
  info: (message, duration = 5000) => Notification.show(message, 'info', duration)
};

// ========================================
// バリデーション
// ========================================

const Validator = {
  /**
   * メールアドレス検証
   */
  email: (email) => {
    const re = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    return re.test(email);
  },
  
  /**
   * 電話番号検証（日本）
   */
  phone: (phone) => {
    const re = /^0\d{9,10}$/;
    return re.test(phone.replace(/-/g, ''));
  },
  
  /**
   * 金額検証
   */
  amount: (amount, min = 0, max = Infinity) => {
    const num = Number(amount);
    return !isNaN(num) && num >= min && num <= max;
  },
  
  /**
   * 必須項目検証
   */
  required: (value) => {
    return value !== null && value !== undefined && value !== '';
  }
};

// ========================================
// 初期化処理
// ========================================

document.addEventListener('DOMContentLoaded', () => {
  // ログイン状態チェック
  const loginInfo = Storage.getLogin();
  
  // ナビゲーションの更新
  updateNavigation(loginInfo);
  
  // ページ固有の初期化
  initializePage();
});

/**
 * ナビゲーション更新
 */
function updateNavigation(loginInfo) {
  const nav = document.querySelector('nav');
  if (!nav) return;
  
  if (loginInfo.userId) {
    // ログイン済み
    const userType = loginInfo.userType;
    const nickname = loginInfo.userData.nickname || loginInfo.userData.name || 'ユーザー';
    
    // マイページリンクを強調
    const mypageLink = nav.querySelector('a[href="mypage.html"]');
    if (mypageLink) {
      mypageLink.innerHTML = `👤 ${nickname}`;
      mypageLink.style.fontWeight = 'bold';
      mypageLink.style.color = '#667eea';
    }
  }
}

/**
 * ページ固有の初期化（各ページで上書き）
 */
function initializePage() {
  // 各ページで実装
}

console.log('✅ わくわくギルド - システム初期化完了');