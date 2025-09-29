import { ArrowRight, Play, Zap } from 'lucide-react';

const Hero = () => {
  return (
    <section className="relative bg-[#E6EFFF] py-20 overflow-hidden">
      <div className="absolute -translate-x-1/2 -translate-y-1/2 top-1/2 left-1/2 w-[150%] h-[150%] rounded-full bg-gradient-radial from-white via-blue-100 to-transparent"></div>

      <div className="relative max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
        <div className="text-center">
          <h1 className="text-4xl md:text-6xl font-bold text-gray-900 mb-4">
            온라인 도구 모음
            <span className="block text-blue-600">77+ Popular Tools</span>
          </h1>

          <p className="text-2xl font-bold text-gray-600 max-w-2xl mx-auto mb-8">
          가장 많이 사용되는 필수 온라인 도구와 AI를 만나보세요.<br />
          누구나 자유롭게 무료로 사용하세요.
        </p>

          <div className="flex justify-center space-x-8 text-gray-700">
            <div className="flex items-center space-x-2">
                <span className="w-3 h-3 bg-blue-400 rounded-full"></span>
                <span className="font-bold">무료 온라인 도구</span>
              </div>
              <div className="flex items-center space-x-2">
                <span className="w-3 h-3 bg-blue-400 rounded-full"></span>
                <span className="font-bold">간편한 교환 형식</span>
              </div>
              <div className="flex items-center space-x-2">
                <span className="w-3 h-3 bg-blue-400 rounded-full"></span>
                <span className="font-bold">누구나 쉬운 기능</span>
              </div>
          </div>
        </div>
      </div>
    </section>
  );
};

export default Hero;