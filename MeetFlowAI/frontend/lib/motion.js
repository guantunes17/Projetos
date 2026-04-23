export const MOTION = {
  duration: {
    quick: 0.2,
    base: 0.28,
    slow: 0.4,
  },
  easeOut: "easeOut",
};

export function fadeInUp(index = 0, reduceMotion = false, options = {}) {
  const baseDelay = options.baseDelay ?? 0.05;
  const y = options.y ?? 12;
  const duration = options.duration ?? MOTION.duration.base;
  return {
    initial: reduceMotion ? { opacity: 1, y: 0 } : { opacity: 0, y },
    animate: { opacity: 1, y: 0 },
    transition: reduceMotion
      ? { duration: 0 }
      : { delay: baseDelay * index, duration, ease: MOTION.easeOut },
  };
}
