import { useNavigate } from 'react-router-dom';

export default function LandingPage() {
  const navigate = useNavigate();

  return (
    <div className="landing">
      <div className="landing__hero">
        <div className="landing__orb">
          <div className="landing__orb-inner" />
        </div>
        <div className="landing__text">
          <h1>
            Welcome to the new <strong>Event Experience Platform</strong>.
          </h1>
          <p>Enjoy the new app to handle your registration for your Deloitte Event.</p>
        </div>
        <button className="btn btn-lg btn-block btn-secondary" onClick={() => navigate('/start')}>
          Start
        </button>
      </div>
      <footer className="footer-disclaimer">
        <p>
          The Event Experience Platform is a new solution for managing participants at Deloitte
          events such as office or department meetings and company runs (e.g., JPMorgan, B2Run). The
          platform was developed by Eike Brenneisen, Andreas Enk and Nils Felten. Currently in its
          pilot phase, the platform may still have a few quirks &ndash; but rest assured, we're
          continuously refining and expanding its capabilities. Over the coming months, it will be
          rolled out to support more and more events across the firm. If you have any questions,
          feedback, or are interested in using the platform for your upcoming event, feel free to get
          in touch with us!
        </p>
      </footer>
    </div>
  );
}
