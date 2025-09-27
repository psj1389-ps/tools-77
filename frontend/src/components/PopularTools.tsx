import { ArrowRight } from 'lucide-react';

const PopularTools = () => {
  const tools = [
    { name: '모든 도구' },
    { name: 'PDF변환도구' },
    { name: '이미지도구' },
    { name: 'AI', featured: true },
    { name: '이미지 편집' },
    { name: '문서도구' },
    { name: 'YOUTUBE' },
    { name: '동영상 편집' },
  ];

  return (
    <section className="py-20 bg-white">
      <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
        <div className="text-center mb-12">
          <h2 className="text-3xl font-bold text-gray-900">인기 도구들</h2>
          <p className="mt-4 text-lg text-gray-600">가장 많이 사용되는 필수 온라인 도구들을 만나보세요.</p>
        </div>
        <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
          {tools.map((tool) => (
            <button
              key={tool.name}
              className="px-6 py-3 rounded-lg font-semibold text-center transition-all duration-200 bg-blue-100 text-blue-800 hover:bg-blue-200 hover:shadow-md"
            >
              {tool.name}
            </button>
          ))}
        </div>
        <div className="text-center mt-12">
          <a href="#" className="text-blue-600 hover:text-blue-800 font-semibold inline-flex items-center">
            77개 이상의 인기 도구가 제공하는 강력한 기능을 확인해보세요.
            <ArrowRight className="ml-2 w-4 h-4" />
          </a>
        </div>
      </div>
    </section>
  );
};

export default PopularTools;