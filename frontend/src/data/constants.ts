import { Tool, Category, FAQ, Feature, Statistic } from '../types';

// 도구 데이터
export const TOOLS: Tool[] = [
  {
    id: 'pdf-converter',
    name: 'PDF 변환기',
    description: '다양한 파일 형식을 PDF로 변환하거나 PDF를 다른 형식으로 변환',
    category: 'converter',
    tags: ['pdf', 'converter', 'document'],
    icon: 'FileText',
    featured: true,
    path: '/tools/pdf-converter'
  },
  {
    id: 'image-resizer',
    name: '이미지 리사이저',
    description: '이미지 크기를 조정하고 최적화',
    category: 'image',
    tags: ['image', 'resize', 'optimize'],
    icon: 'Image',
    featured: true,
    path: '/tools/image-resizer'
  },
  {
    id: 'text-formatter',
    name: '텍스트 포맷터',
    description: '텍스트를 다양한 형식으로 변환하고 정리',
    category: 'text',
    tags: ['text', 'format', 'convert'],
    icon: 'Type',
    featured: true,
    path: '/tools/text-formatter'
  },
  {
    id: 'hash-generator',
    name: '해시 생성기',
    description: 'MD5, SHA1, SHA256 등 다양한 해시값 생성',
    category: 'security',
    tags: ['hash', 'security', 'encrypt'],
    icon: 'Shield',
    featured: true,
    path: '/tools/hash-generator'
  },
  {
    id: 'qr-generator',
    name: 'QR 코드 생성기',
    description: '텍스트나 URL을 QR 코드로 변환',
    category: 'generator',
    tags: ['qr', 'code', 'generator'],
    icon: 'QrCode',
    featured: false,
    path: '/tools/qr-generator'
  },
  {
    id: 'color-picker',
    name: '컬러 피커',
    description: '색상 선택 및 다양한 색상 코드 변환',
    category: 'design',
    tags: ['color', 'picker', 'design'],
    icon: 'Palette',
    featured: false,
    path: '/tools/color-picker'
  }
];

// 카테고리 데이터
export const CATEGORIES: Category[] = [
  {
    id: 'converter',
    name: '변환 도구',
    description: '파일 형식 변환 관련 도구들',
    icon: 'RefreshCw',
    color: 'blue'
  },
  {
    id: 'image',
    name: '이미지 도구',
    description: '이미지 편집 및 처리 도구들',
    icon: 'Image',
    color: 'green'
  },
  {
    id: 'text',
    name: '텍스트 도구',
    description: '텍스트 처리 및 변환 도구들',
    icon: 'Type',
    color: 'purple'
  },
  {
    id: 'security',
    name: '보안 도구',
    description: '암호화 및 보안 관련 도구들',
    icon: 'Shield',
    color: 'red'
  },
  {
    id: 'generator',
    name: '생성 도구',
    description: '다양한 콘텐츠 생성 도구들',
    icon: 'Zap',
    color: 'yellow'
  },
  {
    id: 'design',
    name: '디자인 도구',
    description: '디자인 및 색상 관련 도구들',
    icon: 'Palette',
    color: 'pink'
  }
];

// 특징 데이터
export const FEATURES: Feature[] = [
  {
    id: 'secure',
    title: '🔒 안전한 처리',
    description: '모든 처리는 브라우저에서 로컬로 진행되어 데이터가 안전합니다',
    icon: 'Shield'
  },
  {
    id: 'fast',
    title: '⚡ 빠른 속도',
    description: '로컬 처리로 최적화된 빠른 성능과 속도를 제공합니다',
    icon: 'Zap'
  }
];

// 통계 데이터
export const STATISTICS: Statistic[] = [
  {
    id: 'tools',
    label: '사용 가능한 도구',
    value: '99+',
    description: '다양한 온라인 유틸리티'
  },
  {
    id: 'processing',
    label: '로컬 처리',
    value: '100%',
    description: '브라우저에서 안전하게'
  },
  {
    id: 'access',
    label: '24시간 접근',
    value: '24/7',
    description: '언제든지 사용 가능'
  },
  {
    id: 'cost',
    label: '현재 비용',
    value: '$0',
    description: '완전 무료 서비스'
  }
];

// FAQ 데이터
export const FAQS: FAQ[] = [
  {
    id: 'free',
    question: '정말 무료로 사용할 수 있나요?',
    answer: '네! 저희 도구들은 현재 회원가입 없이 기본 사용에 대해 무료로 제공됩니다. 필수 유틸리티를 접근 가능하게 유지하는 것이 목표입니다.',
    category: 'general'
  },
  {
    id: 'security',
    question: '제 데이터는 얼마나 안전한가요?',
    answer: '저희 도구들은 가능한 한 브라우저에서 로컬로 파일을 처리하도록 설계되었습니다. 개인 데이터를 저장하거나 전송하지 않는 것을 목표로 합니다. 자세한 내용은 개인정보 보호정책을 참조하세요.',
    category: 'security'
  },
  {
    id: 'tools',
    question: '어떤 종류의 도구를 제공하나요?',
    answer: '파일 변환기(PDF, 이미지), 해시 생성기, 텍스트 처리기, 이미지 편집기 등 필수 유틸리티를 제공합니다. 위의 전체 컬렉션을 탐색하거나 소개 페이지에서 자세한 내용을 확인하세요.',
    category: 'tools'
  },
  {
    id: 'account',
    question: '계정을 만들어야 하나요?',
    answer: '계정이 필요하지 않습니다! 어떤 도구 페이지든 방문해서 즉시 사용을 시작하세요. 가능한 한 마찰 없이 도구를 설계했습니다.',
    category: 'general'
  },
  {
    id: 'request',
    question: '새로운 도구나 기능을 요청할 수 있나요?',
    answer: '물론입니다! 사용자 피드백을 바탕으로 유용한 도구를 추가하는 것을 항상 고려하고 있습니다. 제안 사항이 있으시면 연락해 주시면 향후 릴리스에서 고려하겠습니다.',
    category: 'features'
  },
  {
    id: 'filesize',
    question: '어떤 파일 크기가 가장 잘 작동하나요?',
    answer: '저희 도구들은 대부분의 일반적인 파일 크기를 효율적으로 처리하도록 최적화되어 있습니다. 처리가 브라우저에서 이루어지므로 성능은 기기의 성능에 따라 달라지며, 다양한 크기의 파일로 작업할 수 있습니다.',
    category: 'technical'
  }
];