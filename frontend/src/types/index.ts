// 도구 인터페이스
export interface Tool {
  id: string;
  name: string;
  description: string;
  category: string;
  tags: string[];
  icon: string;
  featured: boolean;
  path: string;
}

// 카테고리 인터페이스
export interface Category {
  id: string;
  name: string;
  description: string;
  icon: string;
  color: string;
}

// FAQ 인터페이스
export interface FAQ {
  id: string;
  question: string;
  answer: string;
  category: string;
}

// 특징 인터페이스
export interface Feature {
  id: string;
  title: string;
  description: string;
  icon: string;
}

// 통계 인터페이스
export interface Statistic {
  id: string;
  label: string;
  value: string;
  description: string;
}