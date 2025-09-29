import { Link } from 'react-router-dom';
import { Wrench, Github, Twitter, Mail, Heart, ExternalLink } from 'lucide-react';

const Footer = () => {
  const currentYear = new Date().getFullYear();

  const footerLinks = {
    tools: [
      { name: 'PDF 변환기', href: '/tools/pdf-converter' },
      { name: '이미지 리사이저', href: '/tools/image-resizer' },
      { name: '텍스트 포맷터', href: '/tools/text-formatter' },
      { name: '해시 생성기', href: '/tools/hash-generator' }
    ],
    company: [
      { name: '소개', href: '/about' },
      { name: '블로그', href: '/blog' },
      { name: '문의하기', href: '/contact' },
      { name: '개발자 API', href: '/api' }
    ],
    support: [
      { name: '도움말', href: '/help' },
      { name: 'FAQ', href: '/faq' },
      { name: '개인정보처리방침', href: '/privacy' },
      { name: '이용약관', href: '/terms' }
    ],
    resources: [
      { name: '사용 가이드', href: '/guide' },
      { name: '업데이트', href: '/updates' },
      { name: '피드백', href: '/feedback' },
      { name: '커뮤니티', href: '/community' }
    ]
  };

  const socialLinks = [
    {
      name: 'GitHub',
      href: 'https://github.com',
      icon: <Github className="w-5 h-5" />
    },
    {
      name: 'Twitter',
      href: 'https://twitter.com',
      icon: <Twitter className="w-5 h-5" />
    },
    {
      name: 'Email',
      href: 'mailto:contact@77populartools.com',
      icon: <Mail className="w-5 h-5" />
    }
  ];

  return (
    <footer className="bg-gray-900 text-gray-300">
      {/* 메인 푸터 콘텐츠 */}
      <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-16">
        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-6 gap-8">
          {/* 브랜드 섹션 */}
          <div className="lg:col-span-2">
            <Link to="/" className="flex items-center space-x-2 mb-4">
              <div className="w-12 h-12 rounded-full overflow-hidden">
                <img 
                  src="https://trae-api-sg.mchost.guru/api/ide/v1/text_to_image?prompt=circular%20purple%20logo%20with%20white%20text%20%2277%20TOOLS%22%20modern%20design%20clean%20typography&image_size=square" 
                  alt="77 TOOLS Logo" 
                  className="w-full h-full object-cover"
                />
              </div>
              <span className="text-xl font-bold text-white">
                77+ Popular Tools
              </span>
            </Link>
            
            <p className="text-gray-400 mb-6 max-w-md">
              모든 작업을 위한 완전 무료 온라인 도구 모음. 회원가입 없이 안전하고 빠르게 사용하세요.
            </p>
            
            {/* 소셜 미디어 링크 */}
            <div className="flex space-x-4">
              {socialLinks.map((social) => (
                <a
                  key={social.name}
                  href={social.href}
                  target="_blank"
                  rel="noopener noreferrer"
                  className="p-2 bg-gray-800 hover:bg-blue-600 rounded-lg transition-colors duration-200"
                  aria-label={social.name}
                >
                  {social.icon}
                </a>
              ))}
            </div>
          </div>

          {/* 도구 링크 */}
          <div>
            <h3 className="text-white font-semibold mb-4">인기 도구</h3>
            <ul className="space-y-2">
              {footerLinks.tools.map((link) => (
                <li key={link.name}>
                  <Link
                    to={link.href}
                    className="hover:text-blue-400 transition-colors duration-200"
                  >
                    {link.name}
                  </Link>
                </li>
              ))}
            </ul>
          </div>

          {/* 회사 정보 */}
          <div>
            <h3 className="text-white font-semibold mb-4">회사</h3>
            <ul className="space-y-2">
              {footerLinks.company.map((link) => (
                <li key={link.name}>
                  <Link
                    to={link.href}
                    className="hover:text-blue-400 transition-colors duration-200"
                  >
                    {link.name}
                  </Link>
                </li>
              ))}
            </ul>
          </div>

          {/* 지원 */}
          <div>
            <h3 className="text-white font-semibold mb-4">지원</h3>
            <ul className="space-y-2">
              {footerLinks.support.map((link) => (
                <li key={link.name}>
                  <Link
                    to={link.href}
                    className="hover:text-blue-400 transition-colors duration-200"
                  >
                    {link.name}
                  </Link>
                </li>
              ))}
            </ul>
          </div>

          {/* 리소스 */}
          <div>
            <h3 className="text-white font-semibold mb-4">리소스</h3>
            <ul className="space-y-2">
              {footerLinks.resources.map((link) => (
                <li key={link.name}>
                  <Link
                    to={link.href}
                    className="hover:text-blue-400 transition-colors duration-200"
                  >
                    {link.name}
                  </Link>
                </li>
              ))}
            </ul>
          </div>
        </div>
      </div>

      {/* 하단 저작권 섹션 */}
      <div className="border-t border-gray-800">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-6">
          <div className="flex flex-col md:flex-row justify-between items-center space-y-4 md:space-y-0">
            {/* 저작권 */}
            <div className="flex items-center space-x-2 text-gray-400">
              <span>&copy; {currentYear} 77+ Popular Tools. All rights reserved.</span>
            </div>

            {/* 추가 링크 */}
            <div className="flex items-center space-x-6 text-sm">
              <Link
                to="/privacy"
                className="hover:text-blue-400 transition-colors duration-200"
              >
                개인정보처리방침
              </Link>
              <Link
                to="/terms"
                className="hover:text-blue-400 transition-colors duration-200"
              >
                이용약관
              </Link>
              <Link
                to="/sitemap"
                className="hover:text-blue-400 transition-colors duration-200"
              >
                사이트맵
              </Link>
            </div>


          </div>
        </div>
      </div>

      {/* 맨 위로 버튼 */}
      <button
        onClick={() => window.scrollTo({ top: 0, behavior: 'smooth' })}
        className="fixed bottom-8 right-8 bg-blue-600 hover:bg-blue-700 text-white p-3 rounded-full shadow-lg transition-all duration-200 transform hover:scale-110 z-50"
        aria-label="맨 위로 이동"
      >
        <svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
          <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M5 10l7-7m0 0l7 7m-7-7v18" />
        </svg>
      </button>
    </footer>
  );
};

export default Footer;