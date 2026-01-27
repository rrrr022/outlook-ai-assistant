import React, { useEffect, useState } from 'react';
import { Spinner } from '@fluentui/react-components';
import { useAppStore } from '../store/appStore';
import { outlookService } from '../services/outlookService';
import { aiService } from '../services/aiService';
import { approvalService } from '../services/approvalService';
import { CalendarEvent, TimeSlot } from '../../shared/types';

const CalendarPanel: React.FC = () => {
  const { calendarEvents, setCalendarEvents, settings } = useAppStore();
  const [loading, setLoading] = useState(false);
  const [selectedDate, setSelectedDate] = useState(new Date());
  const [dayPlan, setDayPlan] = useState<string | null>(null);
  const [availableSlots, setAvailableSlots] = useState<TimeSlot[]>([]);

  useEffect(() => {
    loadEvents();
  }, []);

  const loadEvents = async () => {
    setLoading(true);
    try {
      const events = await outlookService.getUpcomingEvents(7);
      setCalendarEvents(events);
      calculateAvailableSlots(events);
    } catch (error) {
      console.error('Error loading events:', error);
    } finally {
      setLoading(false);
    }
  };

  const calculateAvailableSlots = (events: CalendarEvent[]) => {
    // Calculate free time slots for today
    const today = new Date();
    const slots: TimeSlot[] = [];
    const workStart = 9; // 9 AM
    const workEnd = 17; // 5 PM

    for (let hour = workStart; hour < workEnd; hour++) {
      const slotStart = new Date(today);
      slotStart.setHours(hour, 0, 0, 0);
      const slotEnd = new Date(today);
      slotEnd.setHours(hour + 1, 0, 0, 0);

      const conflictingEvent = events.find((event) => {
        const eventStart = new Date(event.start);
        const eventEnd = new Date(event.end);
        return eventStart < slotEnd && eventEnd > slotStart;
      });

      slots.push({
        start: slotStart,
        end: slotEnd,
        available: !conflictingEvent,
        event: conflictingEvent,
      });
    }

    setAvailableSlots(slots);
  };

  const handlePlanDay = async () => {
    setLoading(true);
    try {
      const response = await aiService.chat({
        prompt: `Based on my calendar for today, help me plan my day efficiently. Here are my events: ${JSON.stringify(calendarEvents.slice(0, 5).map(e => ({
          subject: e.subject,
          start: e.start,
          end: e.end,
        })))}. 
        
        Suggest how I should structure my free time and prioritize my work.`,
      });
      setDayPlan(response.content);
    } catch (error) {
      console.error('Error planning day:', error);
    } finally {
      setLoading(false);
    }
  };

  const handleScheduleMeeting = async (slot: TimeSlot) => {
    const meetingSubject = 'New Meeting';
    
    // Request approval before creating meeting
    if (settings.requireApprovalForMeetings) {
      const approved = await approvalService.requestMeetingApproval({
        subject: meetingSubject,
        startTime: slot.start,
        endTime: slot.end,
        source: 'user',
      });
      
      if (approved) {
        loadEvents();
      }
    } else {
      try {
        await outlookService.createCalendarEvent({
          subject: meetingSubject,
          start: slot.start,
          end: slot.end,
        });
        loadEvents();
      } catch (error) {
        console.error('Error creating event:', error);
      }
    }
  };

  const formatTime = (date: Date) => {
    return date.toLocaleTimeString('en-US', { 
      hour: 'numeric', 
      minute: '2-digit',
      hour12: true 
    });
  };

  const formatDate = (date: Date) => {
    return date.toLocaleDateString('en-US', {
      weekday: 'long',
      month: 'short',
      day: 'numeric',
    });
  };

  const todayEvents = calendarEvents.filter((event) => {
    const eventDate = new Date(event.start);
    const today = new Date();
    return eventDate.toDateString() === today.toDateString();
  });

  return (
    <div>
      <h2 className="section-title">📅 Calendar</h2>

      {/* Today's Overview */}
      <div className="card">
        <div className="calendar-header">
          <h3 className="section-header--no-margin">
            {formatDate(new Date())}
          </h3>
          <button 
            className="action-button primary" 
            onClick={handlePlanDay}
            disabled={loading}
          >
            🎯 Plan My Day
          </button>
        </div>

        {loading ? (
          <div className="loading-spinner">
            <Spinner size="small" />
          </div>
        ) : todayEvents.length > 0 ? (
          todayEvents.map((event) => (
            <div key={event.id} className="calendar-slot">
              <span className="slot-time">
                {formatTime(new Date(event.start))}
              </span>
              <span className="slot-event">{event.subject}</span>
            </div>
          ))
        ) : (
          <div className="empty-state">
            <p className="empty-state-text">No events scheduled for today</p>
          </div>
        )}
      </div>

      {/* Day Plan from AI */}
      {dayPlan && (
        <div className="card">
          <h3 className="section-header">
            🎯 Your Day Plan
          </h3>
          <p className="text-medium text-dark pre-wrap">
            {dayPlan}
          </p>
        </div>
      )}

      {/* Available Time Slots */}
      <div className="card mt-12">
        <h3 className="section-header--lg-margin">
          Available Time Slots
        </h3>
        {availableSlots.filter(s => s.available).length > 0 ? (
          availableSlots
            .filter((slot) => slot.available)
            .map((slot, index) => (
              <div 
                key={index} 
                className="calendar-slot calendar-slot--available"
                onClick={() => handleScheduleMeeting(slot)}
              >
                <span className="slot-time">
                  {formatTime(slot.start)}
                </span>
                <span className="slot-event">
                  Available - Click to schedule
                </span>
              </div>
            ))
        ) : (
          <div className="empty-state">
            <p className="empty-state-text">No available slots today</p>
          </div>
        )}
      </div>

      {/* Upcoming Events */}
      <div className="mt-16">
        <h3 className="section-title">Upcoming Events</h3>
        {calendarEvents.slice(0, 5).map((event) => (
          <div key={event.id} className="card">
            <div className="email-subject">{event.subject}</div>
            <div className="email-sender">
              📍 {event.location || 'No location'} · {formatTime(new Date(event.start))}
            </div>
            <div className="email-preview">
              {formatDate(new Date(event.start))}
            </div>
          </div>
        ))}
      </div>
    </div>
  );
};

export default CalendarPanel;
