import * as React from 'react';
import type { ICarouselSlide, ISpfxCarouselProps } from './ISpfxCarouselProps';
import { Swiper, SwiperSlide } from 'swiper/react';
import { Autoplay, FreeMode, Thumbs } from 'swiper/modules';
import type { Swiper as SwiperType } from 'swiper';

import 'swiper/css';
import 'swiper/css/free-mode';
import 'swiper/css/thumbs';

interface ISpfxCarouselState {
  // We track the active slide ourselves so we can draw a custom active state on the thumbnail strip.
  activeIndex: number;
  thumbsSwiper: SwiperType | undefined;
}

function getBackgroundImageStyle(imageSrc: string | undefined, fallback: string): string {
  // SharePoint image URLs can contain parentheses and quotes; wrapping them safely prevents CSS parsing issues.
  if (!imageSrc) {
    return fallback;
  }

  const escapedUrl: string = imageSrc.replace(/"/g, '%22');
  return `url("${escapedUrl}")`;
}

const sectionStyle: React.CSSProperties = {
  overflow: 'hidden',
  color: 'var(--bodyText)'
};

const statusCardStyle: React.CSSProperties = {
  minHeight: '220px',
  display: 'flex',
  alignItems: 'center',
  justifyContent: 'center',
  padding: '32px',
  border: '1px solid #d0d7e2',
  borderRadius: '16px',
  background: '#f8fbff',
  textAlign: 'center'
};

const statusContentStyle: React.CSSProperties = {
  maxWidth: '520px'
};

const statusTitleStyle: React.CSSProperties = {
  margin: '0 0 10px',
  fontSize: '24px',
  lineHeight: 1.2,
  color: '#102033'
};

const statusTextStyle: React.CSSProperties = {
  margin: 0,
  lineHeight: 1.6,
  color: '#605e5c'
};

const rotatorStyle: React.CSSProperties = {
  marginBottom: '5px',
  paddingBottom: '5px'
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

const slidePanelStyle: React.CSSProperties = {
  ...slideLinkStyle,
  cursor: 'default'
};

const imageWrapStyle: React.CSSProperties = {
  position: 'absolute',
  inset: 0,
  width: '100%',
  height: '100%',
  backgroundColor: '#d6dfec',
  backgroundImage: 'linear-gradient(135deg, #dbe6f5 0%, #b8c7dd 55%, #8ea0ba 100%)',
  backgroundPosition: 'center',
  backgroundRepeat: 'no-repeat',
  backgroundSize: 'cover'
};

const overlayStyle: React.CSSProperties = {
  position: 'absolute',
  inset: 0,
  // The main slide keeps a bottom-heavy gradient so text stays readable while most of the image remains visible.
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
  border: '1px solid rgba(15, 23, 42, 0.08)',
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
  // Thumbnails use a stronger solid overlay so very bright images do not wash out the small text.
  background: 'rgba(15, 23, 42, 0.62)'
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
  textShadow: '0 1px 2px rgba(0, 0, 0, 0.45)',
  lineHeight: 1.35
};

const activeThumbIndicatorStyle: React.CSSProperties = {
  position: 'absolute',
  right: '10px',
  bottom: 0,
  left: '10px',
  zIndex: 2,
  height: '4px',
  borderRadius: '999px 999px 0 0',
  background: '#ff6a13',
  boxShadow: '0 0 0 1px rgba(255, 106, 19, 0.12)'
};

export default class SpfxCarousel extends React.Component<ISpfxCarouselProps, ISpfxCarouselState> {
  public constructor(props: ISpfxCarouselProps) {
    super(props);

    this.state = {
      activeIndex: 0,
      thumbsSwiper: undefined
    };
  }

  public render(): React.ReactElement<ISpfxCarouselProps> {
    const {
      slides,
      enableAutoplay,
      autoplayDelaySeconds,
      isLoading,
      errorMessage,
      hasTeamsContext
    } = this.props;
    const { activeIndex, thumbsSwiper } = this.state;

    // Swiper expects milliseconds, but the property pane is friendlier when authors set whole seconds.
    const autoplaySettings = enableAutoplay ? {
      delay: autoplayDelaySeconds * 1000,
      disableOnInteraction: false
    } : false;

    return (
      <section style={sectionStyle} data-teams-context={hasTeamsContext ? 'true' : 'false'}>
        {isLoading ? (
          <div style={statusCardStyle}>
            <div style={statusContentStyle}>
              <h3 style={statusTitleStyle}>Loading carousel items</h3>
              <p style={statusTextStyle}>Fetching the latest rotator items from SharePoint.</p>
            </div>
          </div>
        ) : errorMessage ? (
          <div style={statusCardStyle}>
            <div style={statusContentStyle}>
              <h3 style={statusTitleStyle}>Carousel unavailable</h3>
              <p style={statusTextStyle}>{errorMessage}</p>
            </div>
          </div>
        ) : slides.length === 0 ? (
          <div style={statusCardStyle}>
            <div style={statusContentStyle}>
              <h3 style={statusTitleStyle}>No rotator items found</h3>
              <p style={statusTextStyle}>
                Add items to the News list with a valid target URL and a News Destination of Rotator.
              </p>
            </div>
          </div>
        ) : (
          <>
            <Swiper
              style={rotatorStyle}
              modules={[Autoplay, Thumbs]}
              slidesPerView={1}
              loop={false}
              autoplay={autoplaySettings}
              // realIndex is stable for the visible item and maps cleanly to our thumbnail indicator.
              onSlideChange={(swiper: SwiperType) => this.setState({ activeIndex: swiper.realIndex })}
              thumbs={{
                swiper: thumbsSwiper && !thumbsSwiper.destroyed ? thumbsSwiper : undefined
              }}
            >
              {slides.map((item: ICarouselSlide) => (
                <SwiperSlide key={item.title}>
                  <article style={slideStyle}>
                    {item.href ? (
                      // When authors provide a target URL, the full hero slide becomes the click target.
                      <a style={slideLinkStyle} href={item.href} target="_blank" rel="noreferrer" aria-label={item.title}>
                      <div
                        style={{
                          ...imageWrapStyle,
                          backgroundImage: getBackgroundImageStyle(item.imageSrc, imageWrapStyle.backgroundImage as string)
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
                    ) : (
                      // If no URL exists, we preserve the same visual treatment without rendering a broken link.
                      <div style={slidePanelStyle} aria-label={item.title}>
                        <div
                          style={{
                            ...imageWrapStyle,
                            backgroundImage: getBackgroundImageStyle(item.imageSrc, imageWrapStyle.backgroundImage as string)
                          }}
                          role="img"
                          aria-label={item.imageAlt}
                        />
                        <div style={overlayStyle} aria-hidden="true" />
                        <div style={contentStyle}>
                          <h3 style={slideTitleStyle}>{item.title}</h3>
                          <p style={summaryStyle}>{item.summary}</p>
                        </div>
                      </div>
                    )}
                  </article>
                </SwiperSlide>
              ))}
            </Swiper>
            <Swiper
              style={thumbsStyle}
              modules={[FreeMode, Thumbs]}
              // The thumbs Swiper instance is passed into the main Swiper so the two stay synchronized.
              onSwiper={(swiper: SwiperType) => this.setState({ thumbsSwiper: swiper })}
              spaceBetween={2}
              slidesPerView={Math.min(slides.length, 5)}
              watchSlidesProgress={true}
              freeMode={true}
              loop={false}
              breakpoints={{
                // Keep the thumbnail row dense on desktop but usable on smaller screens.
                0: {
                  slidesPerView: Math.min(slides.length, 2)
                },
                640: {
                  slidesPerView: Math.min(slides.length, 3)
                },
                960: {
                  slidesPerView: Math.min(slides.length, 5)
                }
              }}
            >
              {slides.map((item: ICarouselSlide, index: number) => (
                <SwiperSlide key={`${item.title}-thumb`}>
                  <div style={thumbSlideStyle} aria-hidden="true">
                    <div
                      style={{
                        ...thumbImageStyle,
                        backgroundImage: getBackgroundImageStyle(
                          item.imageSrc,
                          'linear-gradient(135deg, #dbe6f5 0%, #b8c7dd 55%, #8ea0ba 100%)'
                        )
                      }}
                    />
                    <div style={thumbOverlayStyle} />
                    <p style={thumbTitleStyle}>{item.title}</p>
                    {activeIndex === index ? (
                      <div style={activeThumbIndicatorStyle} aria-hidden="true" />
                    ) : null}
                  </div>
                </SwiperSlide>
              ))}
            </Swiper>
          </>
        )}
      </section>
    );
  }
}
