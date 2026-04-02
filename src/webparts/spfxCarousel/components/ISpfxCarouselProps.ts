export interface ICarouselSlide {
  title: string;
  summary: string;
  href?: string;
  imageSrc?: string;
  imageAlt: string;
}

export interface ISpfxCarouselProps {
  slides: ICarouselSlide[];
  enableAutoplay: boolean;
  autoplayDelaySeconds: number;
  isLoading: boolean;
  errorMessage?: string;
  hasTeamsContext: boolean;
}
