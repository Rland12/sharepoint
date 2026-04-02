import * as React from 'react';
import type { ISpfxCarouselProps } from './ISpfxCarouselProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { Swiper, SwiperSlide } from 'swiper/react';
import { Autoplay, FreeMode, Pagination, Thumbs } from 'swiper/modules';
import type { Swiper as SwiperType } from 'swiper';

import 'swiper/css';
import 'swiper/css/free-mode';
import 'swiper/css/pagination';
import 'swiper/css/thumbs';

interface INewsItem {
  category: string;
  title: string;
  summary: string;
  href: string;
  imageSrc: string;
  imageAlt: string;
}

interface ISpfxCarouselState {
  thumbsSwiper: SwiperType | undefined;
}

const DEFAULT_NEWS_ITEMS: INewsItem[] = [
  {
    category: 'Company',
    title: 'The Louie Is an HD Awards Finalist',
    summary: 'The 48-key boutique hotel in Spokane was recognized as a finalist for the 2026 HD Awards guest room category.',
    href: 'https://aka.ms/spfx',
    imageSrc: require('../assets/davenport-hotel.jpg'),
    imageAlt: 'Hotel guest room with layered textures and warm lighting'
  },
  {
    category: 'People',
    title: 'Benefits enrollment reminders open this week',
    summary: 'Review your plan selections, update dependents, and confirm your elections before the deadline.',
    href: 'https://aka.ms/m365pnp',
    imageSrc: require('../assets/pe-survey.jpg'),
    imageAlt: 'Employee reviewing people and benefits information'
  },
  {
    category: 'New Insights',
    title: 'Stephanie Kingsnorth on Adaptive Reuse',
    summary: 'The newspaper spoke with Stephanie Kingsnorth about the renovations planned for the historic Griffith Park Pool in Los Angeles.',
    href: 'https://aka.ms/spfx-yeoman-api',
    imageSrc: require('../assets/riverside-museum.jpg'),
    imageAlt: 'Architectural exterior with strong geometric lines'
  },
  {
    category: 'IT',
    title: 'New secure file sharing process is now available',
    summary: 'Teams can send sensitive documents externally with a simpler approval flow and better audit tracking.',
    href: 'https://aka.ms/spfx-yeoman-api',
    imageSrc: require('../assets/welcome-dark.png'),
    imageAlt: 'Laptop showing a secure collaboration workflow'
  }
];

function getConfiguredSlides(slidesJson: string | undefined): INewsItem[] {
  if (!slidesJson || !slidesJson.trim()) {
    return DEFAULT_NEWS_ITEMS;
  }

  try {
    const parsed: unknown = JSON.parse(slidesJson);

    if (!Array.isArray(parsed)) {
      return DEFAULT_NEWS_ITEMS;
    }

    const slides: INewsItem[] = parsed
      .filter((item: Partial<INewsItem>) =>
        typeof item.category === 'string' &&
        typeof item.title === 'string' &&
        typeof item.summary === 'string' &&
        typeof item.href === 'string' &&
        typeof item.imageSrc === 'string' &&
        typeof item.imageAlt === 'string'
      )
      .map((item: INewsItem) => ({
        category: item.category,
        title: item.title,
        summary: item.summary,
        href: item.href,
        imageSrc: item.imageSrc,
        imageAlt: item.imageAlt
      }));

    return slides.length > 0 ? slides : DEFAULT_NEWS_ITEMS;
  } catch {
    return DEFAULT_NEWS_ITEMS;
  }
}

const sectionStyle: React.CSSProperties = {
  overflow: 'hidden',
  padding: '24px',
  border: '1px solid #e1dfdd',
  borderRadius: '16px',
  background: '#f8fbff',
  color: 'var(--bodyText)'
};

const headerStyle: React.CSSProperties = {
  marginBottom: '16px'
};

const eyebrowStyle: React.CSSProperties = {
  display: 'inline-block',
  marginBottom: '8px',
  color: '#1f5eff',
  fontSize: '12px',
  fontWeight: 700,
  letterSpacing: '0.08em',
  textTransform: 'uppercase'
};

const titleStyle: React.CSSProperties = {
  margin: 0,
  fontSize: '32px',
  lineHeight: 1.15
};

const subtitleStyle: React.CSSProperties = {
  margin: '12px 0 0',
  maxWidth: '540px',
  color: '#605e5c',
  fontSize: '16px',
  lineHeight: 1.5
};

const rotatorStyle: React.CSSProperties = {
  marginBottom: '16px',
  paddingBottom: '24px'
};

const slideStyle: React.CSSProperties = {
  position: 'relative',
  display: 'flex',
  width: '100%'
};

const slideLinkStyle: React.CSSProperties = {
  position: 'relative',
  display: 'flex',
  width: '100%',
  minHeight: '420px',
  padding: '24px',
  overflow: 'hidden',
  border: '1px solid #d0d7e2',
  borderRadius: '16px',
  backgroundColor: '#102033',
  boxSizing: 'border-box',
  textDecoration: 'none'
};

const imageWrapStyle: React.CSSProperties = {
  position: 'absolute',
  inset: 0,
  width: '100%',
  height: '100%',
  backgroundColor: '#e5eefc',
  backgroundPosition: 'center',
  backgroundRepeat: 'no-repeat',
  backgroundSize: 'cover'
};

const overlayStyle: React.CSSProperties = {
  position: 'absolute',
  inset: 0,
  background: 'linear-gradient(180deg, rgba(15, 23, 42, 0) 0%, rgba(15, 23, 42, 0) 48%, rgba(15, 23, 42, 0.34) 68%, rgba(15, 23, 42, 0.76) 100%)'
};

const contentStyle: React.CSSProperties = {
  position: 'relative',
  zIndex: 1,
  display: 'flex',
  flexDirection: 'column',
  justifyContent: 'flex-end',
  minWidth: 0,
  width: '100%',
  maxWidth: '640px',
  marginTop: 'auto'
};

const slideTitleStyle: React.CSSProperties = {
  margin: '0 0 10px',
  fontSize: '30px',
  lineHeight: 1.2,
  color: '#ffffff'
};

const summaryStyle: React.CSSProperties = {
  margin: 0,
  maxWidth: '640px',
  color: 'rgba(255, 255, 255, 0.92)',
  lineHeight: 1.6
};

const thumbsStyle: React.CSSProperties = {
  paddingBottom: '4px'
};

const thumbSlideStyle: React.CSSProperties = {
  position: 'relative',
  height: '92px',
  borderRadius: '12px',
  overflow: 'hidden',
  cursor: 'pointer',
  backgroundColor: '#dbe4f0'
};

const thumbImageStyle: React.CSSProperties = {
  position: 'absolute',
  inset: 0,
  backgroundPosition: 'center',
  backgroundRepeat: 'no-repeat',
  backgroundSize: 'cover'
};

const thumbOverlayStyle: React.CSSProperties = {
  position: 'absolute',
  inset: 0,
  background: 'linear-gradient(180deg, rgba(15, 23, 42, 0.06) 0%, rgba(15, 23, 42, 0.66) 100%)'
};

const thumbTitleStyle: React.CSSProperties = {
  position: 'absolute',
  right: '12px',
  bottom: '10px',
  left: '12px',
  zIndex: 1,
  margin: 0,
  color: '#ffffff',
  fontSize: '12px',
  fontWeight: 700,
  lineHeight: 1.35
};

export default class SpfxCarousel extends React.Component<ISpfxCarouselProps, ISpfxCarouselState> {
  public constructor(props: ISpfxCarouselProps) {
    super(props);

    this.state = {
      thumbsSwiper: undefined
    };
  }

  public render(): React.ReactElement<ISpfxCarouselProps> {
    const {
      description,
      subtitle,
      enableAutoplay,
      autoplayDelay,
      showPagination,
      slidesJson,
      hasTeamsContext
    } = this.props;
    const { thumbsSwiper } = this.state;

    const heading: string = description ? escape(description) : 'News rotator';
    const subheading: string = subtitle ? escape(subtitle) : 'A lightweight rotating news feed built with React and Swiper for SharePoint.';
    const slides: INewsItem[] = getConfiguredSlides(slidesJson);
    const autoplaySettings = enableAutoplay ? {
      delay: autoplayDelay,
      disableOnInteraction: false
    } : false;

    return (
      <section style={sectionStyle} data-teams-context={hasTeamsContext ? 'true' : 'false'}>
        <div style={headerStyle}>
          <span style={eyebrowStyle}>Latest updates</span>
          <h2 style={titleStyle}>{heading}</h2>
          <p style={subtitleStyle}>
            {subheading}
          </p>
        </div>
        <Swiper
          style={rotatorStyle}
          modules={[Autoplay, Pagination, Thumbs]}
          slidesPerView={1}
          loop={slides.length > 1}
          autoplay={autoplaySettings}
          pagination={showPagination ? { clickable: true } : false}
          thumbs={{
            swiper: thumbsSwiper && !thumbsSwiper.destroyed ? thumbsSwiper : undefined
          }}
        >
          {slides.map((item: INewsItem) => (
            <SwiperSlide key={item.title}>
              <article style={slideStyle}>
                <a style={slideLinkStyle} href={item.href} target="_blank" rel="noreferrer" aria-label={item.title}>
                  <div
                    style={{
                      ...imageWrapStyle,
                      backgroundImage: 'url(' + item.imageSrc + ')'
                    }}
                    role="img"
                    aria-label={item.imageAlt}
                  />
                  <div style={overlayStyle} aria-hidden="true" />
                  <div style={contentStyle}>
                    <h3 style={slideTitleStyle}>{item.title}</h3>
                    <p style={summaryStyle}>{item.summary}</p>
                  </div>
                </a>
              </article>
            </SwiperSlide>
          ))}
        </Swiper>
        <Swiper
          style={thumbsStyle}
          modules={[FreeMode, Thumbs]}
          onSwiper={(swiper: SwiperType) => this.setState({ thumbsSwiper: swiper })}
          spaceBetween={12}
          slidesPerView={Math.min(slides.length, 4)}
          watchSlidesProgress={true}
          freeMode={true}
          loop={slides.length > 3}
        >
          {slides.map((item: INewsItem) => (
            <SwiperSlide key={`${item.title}-thumb`}>
              <div style={thumbSlideStyle} aria-hidden="true">
                <div
                  style={{
                    ...thumbImageStyle,
                    backgroundImage: 'url(' + item.imageSrc + ')'
                  }}
                />
                <div style={thumbOverlayStyle} />
                <p style={thumbTitleStyle}>{item.title}</p>
              </div>
            </SwiperSlide>
          ))}
        </Swiper>
      </section>
    );
  }
}
